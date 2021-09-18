import PyQt5.QtWidgets as qt
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QEventLoop, QTimer
import pywinauto
from PyQt5 import uic
import sys
import json
import os

sys.path.append("Modules") # Import modules from Modules directory

from moreWidgets import CheckableComboBox

# =============================================================================
# Notes
# =============================================================================
# Please use pywinauto Version 0.6.3
# This can be insalled with the command pip install pywinauto==0.6.3

def updateJSON():
    """ Updates instances json file """
    with open("data/instances.json", "w") as f:
            json.dump({"Paths": paths, "Instances": instances, "Groups": groups}, f, indent=4)


class Window(qt.QMainWindow):
    
    def __init__(self):
        """ Main Window """
        global instances, paths, groups
        super(Window, self).__init__()
        uic.loadUi("gui/gui.ui", self)
        
        # Get widgets
        createAction = self.findChild(qt.QAction, "createAction") 
        createAction.triggered.connect(self.createWindow)
        openAction = self.findChild(qt.QAction, "openAction") 
        openAction.triggered.connect(self.openUseWin)
        quitAction = self.findChild(qt.QAction, "quitAction") 
        quitAction.triggered.connect(self.close)
        configAction = self.findChild(qt.QAction, "configAction") 
        configAction.triggered.connect(lambda: self.configWin("configPath"))
        locationAction = self.findChild(qt.QAction, "locationAction") 
        locationAction.triggered.connect(lambda: self.configWin("configPath"))
        createGroup = self.findChild(qt.QAction, "createGroup") 
        createGroup.triggered.connect(self.createAutoplayGroup)
        useAction = self.findChild(qt.QAction, "autoPlayMenu") 
        useAction.triggered.connect(self.openUseAction)
        
        
        # Load saved instances
        with open("data/instances.json", "r") as f:
            contents = json.load(f)
            instances = contents["Instances"]
            paths = contents["Paths"]
            groups = contents["Groups"]
            
    def createWindow(self):                                         
        self.w = CreateWindow()
        self.w.show()
        self.hide()
        
    def openUseWin(self):                                         
        self.w = openUseWin()
        self.w.show()
        self.hide()
        
    def createAutoplayGroup(self):
        self.w = createAutoGroup()
        self.w.show()
        self.hide()
        
    def configWin(self, option="configPath"):                               
        self.w = configWin(option)
        self.w.show()
        self.hide()
        
    def openUseAction(self):
        self.w = openUseActionWin()
        self.w.show()
        self.hide()
            
                    
class CreateWindow(qt.QMainWindow):
    
    def __init__(self):
        """ Window for creating instances """
        super(CreateWindow, self).__init__()
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Remove ? window flag

        uic.loadUi("gui/create.ui", self)
        self.color = "#dadada"
        
        # Get widgets
        self.selectBox = self.findChild(qt.QComboBox, "selectBox")
        self.symbolEdit = self.findChild(qt.QLineEdit, "symbolEdit")
        addBtn = self.findChild(qt.QPushButton, "addBtn")
        self.createInsBtn = self.findChild(qt.QPushButton, "createInsBtn")
        self.createBox = self.findChild(qt.QGroupBox, "createBox")
        colorButton = self.findChild(qt.QPushButton, "colorButton")
        self.colorLabel = self.findChild(qt.QLabel, "colorLabel")
        manSelect = self.findChild(qt.QPushButton, "manSelect")
        self.instanceModes = self.findChild(qt.QComboBox, "instanceModes")
        for mode in ["Focus", "Maximise", "Restore"]:
            self.instanceModes.addItem(mode)
        
        self.symbolEdit.textEdited.connect(self.enableInstances)
        
        # Create Checkable ComboBox
        listBox = self.findChild(qt.QGroupBox, "listBox")
        hbox = qt.QHBoxLayout()
        self.windowsList = CheckableComboBox(self)
        self.windowsList.setEnabled(False)
        hbox.addWidget(self.windowsList)
        listBox.setLayout(hbox)
        
        self.newInstance = {}
        
        # Add paths
        for path in sorted(paths):
            self.selectBox.addItem(path)
            
        if paths:
            self.browseInstance()
        
        # Bind Widgets
        self.selectBox.currentTextChanged.connect(self.browseInstance)
        colorButton.clicked.connect(self.changeColor)
        self.createInsBtn.clicked.connect(self.createInstance)
        manSelect.clicked.connect(self.openFileDialog)
        
        self.show()
        
    def enableInstances(self):
        self.manSelect.setEnabled(True)
        self.selectBox.setEnabled(True)
        self.createInsBtn.setEnabled(True)
        
    def closeEvent(self, event):
        """ Run when window gets closed """
        self.w = Window()
        self.w.show()
        self.close()
        
    
    def openFileDialog(self):
        """ Opens file dialog for user to select exe file """
        fname = qt.QFileDialog.getOpenFileName(self, "Select Instance", "C:/",
                                               "EXE files (*.exe);;All files (*.*)")[0]
        self.browseInstance(fname)
        
        
    def browseInstance(self, fname=None):
        """ Open file explorer to select exe file """
        self.addInstance()
        if not fname:
            fname = self.selectBox.currentText()
        self.fname = fname

        # Extract windows
        try:
            app = pywinauto.application.Application().connect(path=self.fname)
            self.dialogs = app.windows(class_name="SCDW_FloatingChart")
            
        except Exception:
            # Show error
            msg = qt.QMessageBox() # Create a message box
            msg.setWindowTitle("Error")
            msg.setText("Application selected is not currently running any additional windows.")
            msg.setIcon(qt.QMessageBox.Critical)
            msg.setStandardButtons(qt.QMessageBox.Ok)
            msg.setDefaultButton(qt.QMessageBox.Ok)
            msg.exec_()
            
        else:
            if self.dialogs:
                # Configure widgets
                self.windowsList.setEnabled(True)
                self.windowsList.clear()
                for dialog in sorted([dialog.window_text() for dialog in self.dialogs]):
                    self.windowsList.addItem(dialog)

                try:
                     previousInstances = self.newInstance[self.selectBox.currentText()]
                     previousInstances = [i[0] for i in previousInstances]
                except KeyError:
                    pass
                else:
                    self.windowsList.checkItems(previousInstances)
                
    
    def changeColor(self):
        """ Opens color picker """
        color = qt.QColorDialog.getColor() # Opens colour picker window
        if color.isValid():
            self.color = color.name()
            self.colorLabel.setStyleSheet(f"background-color:{self.color}; margin:5px; border:2px solid #000000")
                    
                    
    def addInstance(self):
        """ Checks if criteria is met and then adds instance """
        global instances
        symbol = self.symbolEdit.text()
        winNames = self.windowsList.currentData()
        instNames = [[dialog, "SCDW_FloatingChart"] for dialog in winNames]
        
        if symbol and winNames:
            # Add instance
            
            self.newInstance["name"] = symbol
            self.newInstance["color"] = self.color
            self.newInstance[self.fname] = instNames
            
            self.windowsList.clear()
            self.windowsList.setEnabled(False)
            
            
    def createInstance(self):
        """ Write JSON file """
        global instances
        self.addInstance()
        
        symbol = self.newInstance["name"]
        del self.newInstance["name"]
        color = self.newInstance["color"]
        del self.newInstance["color"]
        
        instances[symbol] = {}
        
        for instance in self.newInstance:
            instances[symbol][instance] = self.newInstance[instance]
        instances[symbol]["color"] = color
        instances[symbol]["mode"] = self.instanceModes.currentText()
        
        updateJSON()
            
        self.closeEvent(None)
        
        
class createAutoGroup(qt.QMainWindow):
    
    def __init__(self):
        """ Open window to use saved Instances """
        super(createAutoGroup, self).__init__()
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Remove ? window flag
    
        uic.loadUi("gui/createGroup.ui", self)
        self.isOpen = True
        self.color = "#dadada"
        self.items = []
        
        # Get widgets
        self.selectBox = self.findChild(qt.QComboBox, "selectBox")
        self.nameEdit = self.findChild(qt.QLineEdit, "nameEdit")
        self.createGrpBtn = self.findChild(qt.QPushButton, "createGrpBtn")
        self.createBox = self.findChild(qt.QGroupBox, "createBox")
        colorButton = self.findChild(qt.QPushButton, "colorButton")
        self.colorLabel = self.findChild(qt.QLabel, "colorLabel")
        self.instList = self.findChild(qt.QListWidget, "instList")
        self.delayBox = self.findChild(qt.QSpinBox, "delayBox")
        self.delayBox.setValue(5)
        
        # Bind functions
        self.nameEdit.textEdited.connect(lambda: self.selectBox.setEnabled(True))
        colorButton.clicked.connect(self.changeColor)
        self.selectBox.activated.connect(self.appendButton)
        self.createGrpBtn.clicked.connect(self.addGroup)
        self.instList.clicked.connect(self.removeButton)
        
        # Add values to combobox
        self.addButtons()
    
        self.show()
        
    def closeEvent(self, event):
        """ Run when window gets closed """
        self.isOpen = False
        self.w = Window()
        self.w.show()  
        
    def changeColor(self):
        """ Opens color picker """
        color = qt.QColorDialog.getColor() # Opens colour picker window
        if color.isValid():
            self.color = color.name()
            self.colorLabel.setStyleSheet(f"background-color:{self.color}; margin:5px; border:2px solid #000000")
            
    def addButtons(self):
        """ Adds buttons to combobox """
        global instances
        for instance in instances:
            self.selectBox.addItem(instance)
            
    def appendButton(self):
        """ Adds button to list widget """
        btn = self.selectBox.currentText()
        self.items.append(btn)
        self.instList.clear()
        
        for item in self.items[::-1]:
            self.instList.insertItem(-1, item)
        
        self.createGrpBtn.setEnabled(True)
        
    def addGroup(self):
        """ Adds group to JSON file """
        global groups
        groupName = self.nameEdit.text()
        groupDelay = self.delayBox.value()
        groups[groupName] = {"buttons": self.items, "color": self.color, "delay": groupDelay}
        updateJSON()
        self.close()
        
    def removeButton(self):
        """ Removes button from list widget """
        item = self.instList.currentRow()
        del self.items[item]
        
        self.instList.clear()
        for item in self.items[::-1]:
            self.instList.insertItem(-1, item)
        
        
        
class openUseWin(qt.QMainWindow):
    
    def __init__(self):
        """ Open window to use saved Instances """
        super(openUseWin, self).__init__()
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Remove ? window flag
    
        uic.loadUi("gui/use.ui", self)
        self.isOpen = True
        self.autoPlayOn = False
        
        # Get widgets
        self.scrollArea = self.findChild(qt.QScrollArea, "scrollArea")
        self.deleteBox = self.findChild(qt.QComboBox, "deleteBox")
        deleteBtn = self.findChild(qt.QPushButton, "deleteBtn")
        deleteBtn.clicked.connect(self.deleteInstance)
        
        self.insertButtons()
        self.show()
        
    def closeEvent(self, event):
        """ Run when window gets closed """
        self.isOpen = False
        self.w = Window()
        self.w.show()
        
    def enableAutoPlay(self):
        if not self.autoPlayOn:
            self.autoPlayOn = True
            self.autoBtn.setText("Disable Auto Play")
            self.runAutoPlay()
        else:
            self.autoPlayOn = False
            self.autoBtn.setText("Enable Auto Play")
        
        
    def deleteInstance(self):
        """ Deletes instance """
        global instances
        for instance in instances:
            if instance == self.deleteBox.currentText():
                del instances[instance]
                break
                
        updateJSON()
            
        for button in self.scrollBox.findChildren(qt.QPushButton):
            button.deleteLater()
            
        self.deleteBox.clear()
        self.insertButtons()
    
    
    def insertButtons(self):
        """ Inserts buttons inside of groupbox """
        global instances
        
        # Create layouts
        rowLayout = qt.QHBoxLayout()
        formLayout = qt.QFormLayout()
        self.scrollBox = qt.QGroupBox()
        self.buttonList = []
        rowList, buttonList = [], []
        counter = 4
        
        for instance in instances:
            if counter == 4:
                # Add a new row
                if rowList:
                    formLayout.addRow(rowList[-1])
                rowList.append(qt.QHBoxLayout())
                counter = 0
            
            buttonList.append(qt.QPushButton(instance))
            buttonList[-1].setFont(QFont("Sanserif", 16))
            buttonList[-1].setStyleSheet(f"background-color:{instances[instance]['color']}")
            buttonList[-1].clicked.connect(self.focusWindows)
            rowList[-1].addWidget(buttonList[-1])
            self.deleteBox.addItem(instance)
            self.buttonList.append(buttonList[-1])
            
            counter += 1
        
        if rowList:
            formLayout.addRow(rowList[-1])
                
        self.scrollBox.setLayout(formLayout) # Add form layout to scroll box layout
        self.scrollArea.setWidget(self.scrollBox) # Put groupbox in scroll area
        layout = qt.QVBoxLayout() # Create box layout
        layout.addWidget(self.scrollArea) # Put scroll area in box layout
        
    def focusWindows(self, name=None):
        """ Focus windows of instances """
        global instances
        
        if name:
            instance = instances[name]
        else:
            instance = instances[self.sender().text()]
            
        option = instance["mode"]
    
        for exe in instance:
            if exe != "color" and exe != "mode":
                app = pywinauto.application.Application().connect(path=exe)
            
                # For application
                for win in instance[exe]:
                    # For window
                    try:
                        a = app.window(title=win[0], class_name=win[1])
                        a.set_focus()
                        if option in ["Maximise", "Auto Play and Maximise"]:
                            a.maximize()
                        if option == "Restore":
                            a.restore()
                    except Exception:
                        pass
        self.raise_()
        self.activateWindow()
                        
    def runAutoPlay(self):
        """ Automatically switches between instances """
        counter = 0
        while self.autoPlayOn and self.isOpen:
            loop = QEventLoop()
            self.focusWindows(self.buttonList[counter].text())
            counter += 1
            if counter >= len(self.buttonList):
                counter = 0
            loop = QEventLoop()
            QTimer.singleShot(autoSpeed*1000, loop.quit)
            loop.exec_()
            
class openUseActionWin(qt.QMainWindow):
    
    def __init__(self):
        """ Open window to use saved Groups """
        super(openUseActionWin, self).__init__()
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Remove ? window flag
    
        uic.loadUi("gui/useGroup.ui", self)
        self.isOpen = True
        self.autoPlayOn = False
        
        # Get widgets
        self.scrollArea = self.findChild(qt.QScrollArea, "scrollArea")
        self.deleteBox = self.findChild(qt.QComboBox, "deleteBox")
        self.stopBtn = self.findChild(qt.QPushButton, "stopBtn")
        self.stopBtn.clicked.connect(self.stopAutoplay)
        deleteBtn = self.findChild(qt.QPushButton, "deleteBtn")
        deleteBtn.clicked.connect(self.deleteGroup)
        
        self.insertButtons()
        self.show()
        
    def closeEvent(self, event):
        """ Run when window gets closed """
        self.isOpen = False
        self.w = Window()
        self.w.show()
    
    
    def insertButtons(self):
        """ Inserts buttons inside of groupbox """
        global groups
        
        # Create layouts
        rowLayout = qt.QHBoxLayout()
        formLayout = qt.QFormLayout()
        self.scrollBox = qt.QGroupBox()
        self.buttonList = []
        rowList, buttonList = [], []
        counter = 4
        
        for group in groups:
            if counter == 4:
                # Add a new row
                if rowList:
                    formLayout.addRow(rowList[-1])
                rowList.append(qt.QHBoxLayout())
                counter = 0
            
            buttonList.append(qt.QPushButton(group))
            buttonList[-1].setFont(QFont("Sanserif", 16))
            buttonList[-1].setStyleSheet(f"background-color:{groups[group]['color']}")
            buttonList[-1].clicked.connect(self.focusWindows)
            rowList[-1].addWidget(buttonList[-1])
            self.deleteBox.addItem(group)
            self.buttonList.append(buttonList[-1])
            
            counter += 1
        
        if rowList:
            formLayout.addRow(rowList[-1])
                
        self.scrollBox.setLayout(formLayout) # Add form layout to scroll box layout
        self.scrollArea.setWidget(self.scrollBox) # Put groupbox in scroll area
        layout = qt.QVBoxLayout() # Create box layout
        layout.addWidget(self.scrollArea) # Put scroll area in box layout
        
    def stopAutoplay(self):
        self.autoPlayOn = False
        self.stopBtn.setEnabled(False)
        self.stopBtn.setText("No Group Playing")
        
    def focusWindows(self, name=None):
        """ Focus windows of instances """
        global groups
        
        if name:
            group = groups[name]
        else:
            group = groups[self.sender().text()]
            
        counter = 0
        self.autoPlayOn = True
        self.stopBtn.setEnabled(True)
        self.stopBtn.setText(f"Stop playing {self.sender().text()}")
        
        while self.autoPlayOn and self.isOpen:
            instanceName = group["buttons"][counter]
            loop = QEventLoop()
            
            # Do Action
            instance = instances[instanceName]
            option = instance["mode"]
    
            for exe in instance:
                if exe != "color" and exe != "mode":
                    app = pywinauto.application.Application().connect(path=exe)
                
                    # For application
                    for win in instance[exe]:
                        # For window
                        try:
                            a = app.window(title=win[0], class_name=win[1])
                            a.set_focus()
                            if option == "Maximise":
                                a.maximize()
                            if option == "Restore":
                                a.restore()
                        except Exception:
                            pass
            self.raise_()
            self.activateWindow()
            
            counter += 1
            if counter >= len(group["buttons"]):
                counter = 0
            loop = QEventLoop()
            QTimer.singleShot(group["delay"]*1000, loop.quit)
            loop.exec_()
            
    def deleteGroup(self):
        """ Deletes instance """
        global groups
        for group in groups:
            if group == self.deleteBox.currentText():
                del groups[group]
                break
                
        updateJSON()
            
        for button in self.scrollBox.findChildren(qt.QPushButton):
            button.deleteLater()
            
        self.deleteBox.clear()
        self.insertButtons()
                    
                    
class configWin(qt.QMainWindow):
    
    def __init__(self, option="configPath"):
        """ Open window to use saved Instances """
        super(configWin, self).__init__()
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Remove ? window flag
        
        self.configOptions = {"Saved .exe Locations": "configPath"}
        self.isClose = True
        
        # Call procedure depending on the string passed in
        eval("self." + option + "()")
        
        self.show()
        
    def closeEvent(self, event):
        """ Run when window gets closed """
        if self.isClose:
            self.w = Window()
            self.w.show()
        
    def insertOptions(self, selected):
        """ Inserts config menu options into list widget """
        self.listWidget = self.findChild(qt.QListWidget, "listWidget")
        for option in reversed(self.configOptions):
            self.listWidget.insertItem(-1, option)
        self.listWidget.setCurrentRow(list(self.configOptions.keys()).index(selected))
        self.listWidget.clicked.connect(self.listWidget_clicked)
        
        
    def listWidget_clicked(self):
        """ Calls function depending on selected item """
        item = self.listWidget.currentItem().text()
        self.isClose = False
        self.close()
        self.w = configWin(self.configOptions[item])
        self.isClose = True
        
        
    def deleteItem(self):
        """ Deletes path """
        item = self.pathList.currentItem()
    
        msg = qt.QMessageBox() 
        msg.setWindowTitle("Are you Sure?")
        msg.setText(f"Are you sure you want to delete '{item.text()}' ?") 
        msg.setIcon(qt.QMessageBox.Question)
        msg.setStandardButtons(qt.QMessageBox.Yes|qt.QMessageBox.No) 
        msg.setDefaultButton(qt.QMessageBox.No) 
        msg.buttonClicked.connect(self.confirmDelete)
        x = msg.exec_()
        
        
    def confirmDelete(self, i):
        global paths
        option = i.text()[1:]
        item = self.pathList.currentItem().text()
        
        if option == "Yes":
            paths.remove(item)
            self.insertFilePaths()
            updateJSON()
            
    
    def addNewPath(self):
        global paths
        """ Adds a new exe path """
        fname = qt.QFileDialog.getOpenFileName(self, "Select Instance", "C:/",
                                               "EXE files (*.exe);;All files (*.*)")[0]
        try:
            fname
        except NameError:
            pass
        else:
            # Add path
            if fname:
                paths.append(fname)
                self.insertFilePaths()
                updateJSON()
                
    def insertFilePaths(self):
        """ Inserts filepaths into list widget """
        self.pathList.clear()
        for path in reversed(sorted(paths)):
            self.pathList.insertItem(-1, path)
        
            
    def configPath(self):
        """ Opens config path selection """
        uic.loadUi("gui/configPath.ui", self)
        self.insertOptions("Saved .exe Locations")
        
        # Load widgets
        addPath = self.findChild(qt.QPushButton, "addPath") 
        addPath.clicked.connect(self.addNewPath)
        self.pathList = self.findChild(qt.QListWidget, "pathList")
        self.pathList.clicked.connect(self.deleteItem)
        
        self.insertFilePaths()
            

if __name__ == "__main__":
    if not os.path.isfile("data/instances.json"):
         with open("data/instances.json", "w") as f:
            json.dump({"Paths": [], "Groups": {}, "Instances": {}}, f, indent=4)
    
    app = qt.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())

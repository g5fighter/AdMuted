# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QAction, QMenu, QSystemTrayIcon

#mi codigo
#import ctypes
#import wmi
import win32api, win32con, win32process, win32gui
import sched, time
from pycaw.pycaw import AudioUtilities
from PyQt5.QtCore import QTimer
import os.path
import json
import subprocess

filename = "Spotify.exe"
spotifyPath = ""
processpid = -1
actual_song = ""
is_muting = False
open_onStart = False
hide_onStart = False
s = sched.scheduler(time.time, time.sleep)
exit = 0
ui = ""
datafilename = ".data"

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def getSpotifyData():
    processes = win32process.EnumProcesses()    # get PID list
    for pid in processes:
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
            exe = win32process.GetModuleFileNameEx(handle, 0)
            if filename.lower() in exe.lower():
                global spotifyPath
                spotifyPath = exe
                global processpid
                processpid = pid
                return True
        except:
            pass
    return False

def winEnumHandler( hwnd, ctx):
    if win32gui.IsWindowVisible( hwnd ):
        if( win32process.GetWindowThreadProcessId(hwnd)[1] == processpid):
            global actual_song
            actual_song = win32gui.GetWindowText( hwnd )

def getWindowName():
    win32gui.EnumWindows(winEnumHandler, None)

def mute_app():
    if is_muting:
        getWindowName()
        ui.update_song(actual_song)
        if "Advertisement" in actual_song:  #Spotify app is named as Advertisement
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                volume = session.SimpleAudioVolume
                volume.SetMute(1, None)
        elif "Spotify" in actual_song:      #App named as Spotify(Only when ad plays, else it's Spotify Free)
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                volume = session.SimpleAudioVolume
                volume.SetMute(1, None)
        else:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                volume = session.SimpleAudioVolume
                volume.SetMute(0, None)


def initialConf():
    if os.path.isfile(datafilename):
        with open(datafilename) as f:
            data = json.load(f)
        global spotifyPath
        spotifyPath = data['SpotifyPath']
        global open_onStart
        open_onStart = bool(data['OpenStart'])
        global hide_onStart
        hide_onStart = bool(data['HideStart'])
    else:
        data = {}
        data['SpotifyPath']= spotifyPath
        data['OpenStart']=open_onStart
        data['HideStart']=hide_onStart
        with open(datafilename, 'w') as outfile:
            json.dump(data, outfile, indent=2)
    ui.update_checkBox()

def writeInJson(etiqueta, dato):
    with open(datafilename, 'r+') as f:
        data = json.load(f)
        data[etiqueta] = dato # <--- add `id` value.
        f.seek(0)        # <--- should reset file position to the beginning.
        json.dump(data, f, indent=2)
        f.truncate()     # remove remaining part
## FIN CODIGO PROPIO

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(416, 193)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(416, 193))
        MainWindow.setMaximumSize(QtCore.QSize(416, 193))
        MainWindow.setMouseTracking(False)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(resource_path("iconadmuted.png")), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())
        self.centralwidget.setSizePolicy(sizePolicy)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(-1, -1, 449, 171))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout.setContentsMargins(5, 5, 5, 5)
        self.verticalLayout.setObjectName("verticalLayout")
        self.TExt = QtWidgets.QLabel(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.TExt.sizePolicy().hasHeightForWidth())
        self.TExt.setSizePolicy(sizePolicy)
        self.TExt.setMinimumSize(QtCore.QSize(0, 20))
        self.TExt.setMaximumSize(QtCore.QSize(16777215, 16777215))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.TExt.setFont(font)
        self.TExt.setAutoFillBackground(False)
        self.TExt.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignTop)
        self.TExt.setObjectName("TExt")
        self.verticalLayout.addWidget(self.TExt)
        self.MuteButton = QtWidgets.QPushButton(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.MuteButton.sizePolicy().hasHeightForWidth())
        self.MuteButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(15)
        self.MuteButton.setFont(font)
        self.MuteButton.setObjectName("MuteButton")
        self.verticalLayout.addWidget(self.MuteButton)
        self.OpenSpotifyCheck = QtWidgets.QCheckBox(self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.OpenSpotifyCheck.setFont(font)
        self.OpenSpotifyCheck.setObjectName("OpenSpotifyCheck")
        self.verticalLayout.addWidget(self.OpenSpotifyCheck)
        self.HideAppCheck = QtWidgets.QCheckBox(self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.HideAppCheck.setFont(font)
        self.HideAppCheck.setObjectName("HideAppCheck")
        self.verticalLayout.addWidget(self.HideAppCheck)
        self.PlayingNowLabel = QtWidgets.QLabel(self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.PlayingNowLabel.sizePolicy().hasHeightForWidth())
        self.PlayingNowLabel.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.PlayingNowLabel.setFont(font)
        self.PlayingNowLabel.setAlignment(QtCore.Qt.AlignBottom|QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft)
        self.PlayingNowLabel.setObjectName("PlayingNowLabel")
        self.verticalLayout.addWidget(self.PlayingNowLabel)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setEnabled(True)
        self.statusbar.setToolTip("")
        self.statusbar.setStatusTip("")
        self.statusbar.setAccessibleName("")
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "AdMuted - An spotify ad silencer"))
        self.TExt.setText(_translate("MainWindow", "Start enjoying and no more adSuffering"))
        self.MuteButton.setText(_translate("MainWindow", "Start Muting"))
        self.OpenSpotifyCheck.setText(_translate("MainWindow", "Open Spotify on Start Muting"))
        self.HideAppCheck.setText(_translate("MainWindow", "Hide app on Start Muting"))
        self.PlayingNowLabel.setText(_translate("MainWindow", "Playing now:"))

        #MI CÓDIGO
        self.MuteButton.clicked.connect(self.muteButton)
        self.OpenSpotifyCheck.clicked.connect(self.open_at_start)
        self.HideAppCheck.clicked.connect(self.hide_at_start)

    def open_at_start(self):
        global open_onStart
        open_onStart = not open_onStart
        writeInJson('OpenStart', open_onStart)
        
    def hide_at_start(self):
        global hide_onStart
        hide_onStart = not hide_onStart
        writeInJson('HideStart', hide_onStart)

    def showWindow(self):
        MainWindow.show()   
        self.tray.hide()   

    def convertInTray(self):
        MainWindow.hide()
        ## MI COdigo
        # Create the tray
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(QtGui.QIcon(resource_path("iconadmuted.png")))
        self.tray.show()

        # Create the menu
        self.menu = QMenu()

        self.action1 = QAction("Return")
        self.action1.triggered.connect(self.showWindow)
        self.menu.addAction(self.action1)

        self.quit = QAction("Quit")
        self.quit.triggered.connect(app.quit)
        self.menu.addAction(self.quit)
        ##
        self.tray.setContextMenu(self.menu)



    
    def muteButton(self):
        global is_muting
        if not is_muting:
            if processpid==-1:
                if open_onStart:
                    try:
                        subprocess.call([spotifyPath])
                        getSpotifyData()
                        is_muting = not is_muting
                        self.MuteButton.setText("Is Muting")
                    except:
                        self.update_song("[Error]: Cannot open Spotify please do it manually")
                        return
                else:
                    self.PlayingNowLabel.setText("Spotify is not running, open it and try again")  
            else:
                    is_muting = not is_muting
                    self.MuteButton.setText("Is Muting")
            if hide_onStart:
                self.convertInTray()
        else:
            is_muting = not is_muting
            self.MuteButton.setText("Start Muting")   
            ui.update_song("")         
    
    def update_checkBox(self):
        self.OpenSpotifyCheck.setChecked(open_onStart)
        self.HideAppCheck.setChecked(hide_onStart)

    def update_song(self, song=actual_song):
        self.PlayingNowLabel.setText("Playing now: "+song)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    getSpotifyData()
    initialConf()
    timer = QTimer()
    timer.timeout.connect(mute_app)
    timer.start(1000)
    sys.exit(app.exec_())
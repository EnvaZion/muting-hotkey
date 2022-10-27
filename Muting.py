from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QShortcut, QApplication, QSystemTrayIcon, QStyle, QAction, qApp, QMenu, QDialog
from pycaw.utils import AudioUtilities
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import QAbstractNativeEventFilter, QAbstractEventDispatcher, QStandardPaths, Qt
from pyqtkeybind import keybinder
import os, sys
from BlurWindow.blurWindow import GlobalBlur
from win32api import GetMonitorInfo, MonitorFromPoint



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


icon_dir = resource_path("Icons")


class WinEventFilter(QAbstractNativeEventFilter):
    def __init__(self, keybinder):
        self.keybinder = keybinder
        super().__init__()

    def nativeEventFilter(self, eventType, message):
        ret = self.keybinder.handler(eventType, message)
        return ret, 0


class HotkeyDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Hotkey settings")
        self.resize(225, 225)
        self.move(800, 400)
        self.Inputline = QtWidgets.QLineEdit(self)
        self.Inputline.setGeometry(QtCore.QRect(105, 62, 100, 20))


        self.Ctrl = QtWidgets.QRadioButton('CTRL', self)
        self.Ctrl.setGeometry(QtCore.QRect(20, 30, 80, 20))
        self.Ctrl.setChecked(True)

        self.Shift = QtWidgets.QRadioButton('SHIFT', self)
        self.Shift.setGeometry(QtCore.QRect(20, 60, 80, 20))

        self.Alt = QtWidgets.QRadioButton('ALT', self)
        self.Alt.setGeometry(QtCore.QRect(20, 90, 80, 20))

        self.CustomHotkey = QtWidgets.QRadioButton('Custom:', self)
        self.CustomHotkey.setGeometry(QtCore.QRect(20, 130, 80, 20))

        self.CustomInputline = QtWidgets.QLineEdit(self)
        self.CustomInputline.setGeometry(QtCore.QRect(105, 130, 100, 20))


        self.Plus_label = QtWidgets.QLabel('+', self)
        self.Plus_label.setGeometry(QtCore.QRect(80, 65, 10, 10))

        self.buttonBox = QtWidgets.QDialogButtonBox(self)
        self.buttonBox.setGeometry(QtCore.QRect(50, 190, 156, 23))

        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.accepted.connect(self.ok_pressed)
        self.buttonBox.rejected.connect(self.cancel_pressed)

        # self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)

    def ok_pressed(self):

        if self.Ctrl.isChecked() == True:
            Modifier = 'Ctrl'
            Hotkey = Modifier + '+' + str(self.Inputline.text())
        if self.Shift.isChecked() == True:
            Modifier = 'Shift'
            Hotkey = Modifier + '+' + str(self.Inputline.text())
        if self.Alt.isChecked() == True:
            Modifier = 'Alt'
            Hotkey = Modifier + '+' + str(self.Inputline.text())
        if self.CustomHotkey.isChecked() == True:
            if str(self.CustomInputline.text()) != '':
                Hotkey = str(self.CustomInputline.text())
                Hotkey = ''.join(Hotkey.split())
            else:
                HotkeyDialog.cancel_pressed(self)
                return



        settings = QtCore.QSettings('Olegs Apps', 'Muting Hotkey')
        PreviousHotkey = settings.value('Hotkey')
        keybinder.init()
        keybinder.unregister_hotkey(w.winId(), PreviousHotkey)

        settings.setValue('Hotkey', Hotkey)
        keybinder.register_hotkey(w.winId(), Hotkey, w.gotovar)
        win_event_filter = WinEventFilter(keybinder)
        event_dispatcher = QAbstractEventDispatcher.instance()
        event_dispatcher.installNativeEventFilter(win_event_filter)
        self.hide()
        w.show()


    def cancel_pressed(self):
        self.hide()
        w.show()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        w.show()


def muteapp(appname):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == appname:
            if volume.GetMute() == 0:
                volume.SetMute(1, None)
                return print(appname, 'muted')

            else:
                volume.SetMute(0, None)
                return print(appname, 'unmuted')


def setvol(appname, decibels):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        interface = session.SimpleAudioVolume
        if session.Process and session.Process.name() == appname:
            # only set volume in the range 0.0 to 1.0
            volume = min(1.0, max(0.0, decibels))
            interface.SetMasterVolume(volume, None)
            print("Volume set to", volume)  # debug


def getvol(appname):
    if appname != '':

        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == appname:
                return float("{0:.2f}".format(interface.GetMasterVolume()))
    else:
        return 0


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.resize(307, 95)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        # Window settings
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)  # frameless and alt-tabless
        app.focusChanged.connect(self.on_focusChanged)

        GlobalBlur(self.winId(), Acrylic=True, Dark=False)

        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(23, 21, 141, 22))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.setStyleSheet("font-family:Arial; font-size: 15px;")

        # Combobox filling without 'Process: '
        self.comboBox_Fill()

        # Buttons

        self.button = QtWidgets.QPushButton(self.centralwidget)
        self.button.setGeometry(QtCore.QRect(203, 20, 24, 24))
        self.button.clicked.connect(self.gotovar)
        self.button.setIcon(QtGui.QIcon(resource_path('Icons/mute.png')))

        self.button1 = QtWidgets.QPushButton(self.centralwidget)
        self.button1.setGeometry(QtCore.QRect(173, 20, 24, 24))
        self.button1.clicked.connect(self.comboBox_Fill)
        self.button1.setIcon(QtGui.QIcon(resource_path('Icons/refresh.png')))

        self.button2 = QtWidgets.QPushButton(self.centralwidget)
        self.button2.setGeometry(QtCore.QRect(233, 20, 24, 24))
        self.button2.clicked.connect(self.CMH_pressed)
        self.button2.setIcon(QtGui.QIcon(resource_path('Icons/settings.png')))

        self.button3 = QtWidgets.QPushButton(self.centralwidget)
        self.button3.setGeometry(QtCore.QRect(263, 20, 24, 24))
        self.button3.clicked.connect(QApplication.instance().quit)
        self.button3.setIcon(QtGui.QIcon(resource_path('Icons/power.png')))

        # Slider
        self.volslider = QtWidgets.QSlider(self.centralwidget)
        self.volslider.setGeometry(QtCore.QRect(23, 55, 263, 23))
        self.volslider.setOrientation(QtCore.Qt.Horizontal)
        self.volslider.setMinimum(1)
        self.volslider.setMaximum(100)
        self.volslider.setSingleStep(1)
        self.changeSliderValue()
        self.volslider.valueChanged.connect(self.volumeCont)
        self.comboBox.currentTextChanged.connect(self.changeSliderValue)
        self.volslider.setTickPosition(2)
        # self.volslider.setTickInterval(5)


        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        self.statusbar.setSizeGripEnabled(0)
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.quitSc = QShortcut(QKeySequence('Ctrl+Q'), self)
        self.quitSc.activated.connect(QApplication.instance().quit)
        self.icon = QSystemTrayIcon(self)

        self.icon.activated[QSystemTrayIcon.ActivationReason].connect(self.click_show)

        # Tray menu
        self.icon.setIcon(QtGui.QIcon(resource_path('Icons/soundon.png')))
        self.setWindowIcon(QtGui.QIcon(resource_path('Icons/soundon.png')))

        show_action = QAction("Show", self)
        quit_action = QAction("Exit", self)
        hide_action = QAction("Hide", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(qApp.quit)
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.icon.setContextMenu(tray_menu)
        self.icon.show()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Muting Hotkey"))

    def closeEvent(self, event):
        event.ignore()
        self.hide()


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    x = 0

    def gotovar(self):
        appname = self.comboBox.currentText()
        muteapp(appname=appname)
        if MainWindow.x == 0:
            self.icon.setIcon(QtGui.QIcon(resource_path('Icons/soundoff.png')))
            MainWindow.x = 1
            return appname

        if MainWindow.x == 1:
            self.icon.setIcon(QtGui.QIcon(resource_path('Icons/soundon.png')))
            MainWindow.x = 0
            return appname

    def volumeCont(self):
        appname = self.comboBox.currentText()
        decibels = self.volslider.value() / 100
        setvol(appname=appname, decibels=decibels)

    def comboBox_Fill(self):
        self.comboBox.clear()
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            session = str(session)
            if 'Process' in session:
                session = session[9:]
                self.comboBox.addItem(session)

    def changeSliderValue(self):
        appname = self.comboBox.currentText()
        value = int(getvol(appname=appname) * 100)

        self.volslider.setValue(value)

    def click_show(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            if self.isHidden():
                self.show()
                self.activateWindow()
                if self.comboBox.count() == 0:
                    print('0')
                    self.comboBox_Fill()
            elif self.isVisible():
                self.hide()

    def on_focusChanged(self):
        if self.isActiveWindow() == False and HotkeyDialog(self).isVisible() == True:
            w.hide()
        if self.isActiveWindow() == False and HotkeyDialog(self).isVisible() == False:
            w.hide()

    def CMH_pressed(self):
        HotkeyDialog(self).exec()


if __name__ == "__main__":
    import sys

    #display settings
    monitor_info = GetMonitorInfo(MonitorFromPoint((0, 0)))
    monitor_area = monitor_info.get("Monitor")
    work_area = monitor_info.get("Work")
    taskbarheight = monitor_area[3] - work_area[3]


    keybinder.init()
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    settings = QtCore.QSettings('Olegs Apps', 'Muting Hotkey')
    mutinghotkey = settings.value('Hotkey')

    # first launch settings
    if mutinghotkey == None or mutinghotkey == '':
        settings.setValue('Hotkey', 'Shift+PgDown')
        mutinghotkey = settings.value('Hotkey')

    keybinder.register_hotkey(w.winId(), mutinghotkey, w.gotovar)
    win_event_filter = WinEventFilter(keybinder)
    event_dispatcher = QAbstractEventDispatcher.instance()
    event_dispatcher.installNativeEventFilter(win_event_filter)

    w.show()
    w.activateWindow()
    w.move(monitor_area[2]-307, monitor_area[3]-taskbarheight-95)
    sys.exit(app.exec())

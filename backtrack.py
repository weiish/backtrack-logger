import wx
import wx.adv
import os.path
from activity import Activity
import configparser
import time
from datetime import datetime, timedelta
import psutil
import win32gui
import win32process

TRAY_TOOLTIP = 'Backtrack Data'
TRAY_ICON = 'icon.png'

def create_menu_item(menu, label, func):
    item = wx.MenuItem(menu, -1, label)
    menu.Bind(wx.EVT_MENU, func, id=item.GetId())
    menu.Append(item)
    return item

class MainFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        super(MainFrame, self).__init__(*args, **kwargs)
        self.timerDelay = 1000
        self.loadSaveFolder()
        self.InitUI()
        
    def InitUI(self):
        #MENU BAR
        menuBar = wx.MenuBar()
        fileMenu = wx.Menu()
        fileItem = fileMenu.Append(wx.ID_EXIT, 'Quit', 'Exit application')
        menuBar.Append(fileMenu, '&File')
        self.SetMenuBar(menuBar)

        self.Bind(wx.EVT_MENU, self.onQuit, fileItem)
        self.Bind(wx.EVT_ICONIZE, self.onMinimize)
        self.Bind(wx.EVT_CLOSE, self.onClose)

        self.SetSize((300, 200))
        self.SetTitle('Backtrack Data Collector (Minimizes to Tray)')
        self.Centre()

        #TIMER
        self.appStartTime = datetime.now()
        self.processTimeActive = 0
        self.timer= wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.updateTimer, self.timer)
        self.lastWindowName = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        self.pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1]
        self.processName = psutil.Process(self.pid).name()
        self.startTime = time.time()
        self.timer.Start(self.timerDelay)

        #PANEL
        panel = wx.Panel(self)
        sizer = wx.GridBagSizer(15, 15)
        boldfont = wx.Font(10, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        text1 = wx.StaticText(panel, label="Log Folder:")
        text1.SetFont(boldfont)
        sizer.Add(text1, pos=(5,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.filePathText = wx.StaticText(panel, label=self.filePath)
        self.filePathText.SetFont(font)
        sizer.Add(self.filePathText, pos=(5,1), span=(1,3), flag=wx.TOP|wx.LEFT|wx.BOTTOM|wx.RIGHT|wx.EXPAND, border=10)

        btn1 = wx.Button(panel, label="Change Folder")
        btn1.SetFont(font)
        sizer.Add(btn1, pos=(5,5), flag=wx.TOP|wx.RIGHT|wx.LEFT|wx.BOTTOM, border=5)
        btn1.Bind(wx.EVT_BUTTON, self.on_change_folder)

        #total time active
        statuslabeltext = wx.StaticText(panel, label='Status: ')
        statuslabeltext.SetFont(boldfont)
        sizer.Add(statuslabeltext, pos=(0,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.statustext = wx.StaticText(panel, label='Logging for ' + str(datetime.now() - self.appStartTime)[:-3])
        self.statustext.SetFont(font)
        sizer.Add(self.statustext, pos=(0,1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        #pid
        pidLabel = wx.StaticText(panel, label='PID: ')
        pidLabel.SetFont(boldfont)
        sizer.Add(pidLabel, pos=(1,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.pidText = wx.StaticText(panel, label=str(self.pid))
        self.pidText.SetFont(font)
        sizer.Add(self.pidText, pos=(1,1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        #processName
        pnameLabel = wx.StaticText(panel, label='Process Name: ')
        pnameLabel.SetFont(boldfont)
        sizer.Add(pnameLabel, pos=(2,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.pnameText = wx.StaticText(panel, label=self.processName)
        self.pnameText.SetFont(font)
        sizer.Add(self.pnameText, pos=(2,1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        #windowName
        wnameLabel = wx.StaticText(panel, label='Window Name: ')
        wnameLabel.SetFont(boldfont)
        sizer.Add(wnameLabel, pos=(3,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.wnameText = wx.StaticText(panel, label=self.lastWindowName)
        self.wnameText.SetFont(font)
        sizer.Add(self.wnameText, pos=(3,1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        #time active
        timeLabel = wx.StaticText(panel, label='Current Active Time: ')
        timeLabel.SetFont(boldfont)
        sizer.Add(timeLabel, pos=(4,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        self.timeText = wx.StaticText(panel, label=time.strftime('%H:%M:%S', time.gmtime(self.processTimeActive)))
        self.timeText.SetFont(font)
        sizer.Add(self.timeText, pos=(4,1), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=10)

        panel.SetSizer(sizer)
        sizer.Fit(self)
    
    def loadSaveFolder(self):
        configFilePath = 'config.ini'
        config = configparser.ConfigParser()
        config.read(configFilePath)
        config['DEFAULT']['path'] = os.path.dirname(os.path.abspath(__file__)) + '\\logs'
        if config.has_option('SETTINGS','path'):
            self.filePath = config['SETTINGS']['path']
        else:
            self.filePath = config['DEFAULT']['path']
        with open(configFilePath, 'w') as configFile:
            config.write(configFile)
        return True
    
    def updateConfig(self, newPath):
        configFilePath = 'config.ini'
        config = configparser.ConfigParser()
        config.read(configFilePath)
        if not config.has_section('SETTINGS'):
            config.add_section('SETTINGS')
        config['SETTINGS']['path'] = newPath
        with open(configFilePath, 'w') as configFile:
            config.write(configFile)
        self.filePath=newPath
        self.filePathText.SetLabel(newPath)
        return True

    def updateGUI(self):
         #total time active
        self.statustext.SetLabel('Logging for ' + str(datetime.now() - self.appStartTime)[:-3])
        #pid
        self.pidText.SetLabel(str(self.pid))
        #processName
        self.pnameText.SetLabel(self.processName)
        #windowName
        self.wnameText.SetLabel(self.lastWindowName)
        #time active
        self.timeText.SetLabel(time.strftime('%H:%M:%S', time.gmtime(self.processTimeActive)))
        self.Update()

    def on_change_folder(self, event):
        print('Changing folders!')
        dlg = wx.DirDialog(self, message="Choose a directory:", defaultPath=self.filePath)
        if dlg.ShowModal() == wx.ID_OK:
            self.updateConfig(dlg.GetPath())
            self.Update()
        dlg.Destroy()

    def updateTimer(self, event):
        self.processTimeActive += self.timerDelay / 1000
        windowName = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if self.processName == "ERROR-FINDING-PROCESS":
            try:
                self.pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1]
                self.processName = psutil.Process(self.pid).name()
            except:
                self.processName = "ERROR-FINDING-PROCESS"
        if self.lastWindowName != windowName:
            #OutputActivity(startTime, time.time(), lastWindowName, processName)
            activity = Activity(self.startTime, time.time(), self.lastWindowName, self.processName, self.filePath)
            activity.SaveToFile()
            self.startTime = time.time()
            try:
                self.pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[-1]
                self.processName = psutil.Process(self.pid).name()
            except:
                self.processName = "ERROR-FINDING-PROCESS"
            self.lastWindowName = windowName
            del activity
            self.processTimeActive = 0
        self.updateGUI()
        

    def onQuit(self, e):
        self.timer.Stop()
        self.Close()
    
    def onClose(self, evt):
        self.Hide() 
    
    def onMinimize(self, evt):
        if self.IsIconized():
            self.Hide()

class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        self.frame = frame
        super(TaskBarIcon, self).__init__()
        self.set_icon(TRAY_ICON)
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.on_left_down)

    def CreatePopupMenu(self):
        menu = wx.Menu()
        create_menu_item(menu, 'Change Data Folder', self.on_change_folder)        
        menu.AppendSeparator()
        create_menu_item(menu, 'Show Menu', self.on_show_menu)
        menu.AppendSeparator()
        create_menu_item(menu, 'Exit Backtrack Data Collector', self.on_exit)
        return menu

    def set_icon(self, path):
        icon = wx.Icon(wx.Bitmap(path))
        self.SetIcon(icon, TRAY_TOOLTIP)

    def on_left_down(self, event):
        self.on_show_menu(event)

    def on_change_folder(self, event):
        dlg = wx.DirDialog(self.frame, message="Choose a directory:", defaultPath=self.frame.filePath)
        if dlg.ShowModal() == wx.ID_OK:
            self.frame.updateConfig(dlg.GetPath())
            self.frame.Update()
        dlg.Destroy()

    def on_show_menu(self, event):
        self.frame.Show()
        self.frame.Restore()

    def on_exit(self, event):
        wx.CallAfter(self.Destroy)
        self.frame.Destroy()

class App(wx.App):
    def OnInit(self):
        frame=MainFrame(None, style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)
        frame.SetIcon(wx.Icon(TRAY_ICON))
        self.SetTopWindow(frame)
        TaskBarIcon(frame)
        frame.Show()
        return True

def main():
    app = App(False)
    app.MainLoop()


if __name__ == '__main__':
    main()
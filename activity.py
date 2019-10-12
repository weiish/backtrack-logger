import datetime
import csv
import sys
import time
import psutil
import win32gui
import win32process
import os

class Activity:
    def __init__(self, startTime=0, endTime=0, windowName='', processName='', path=''):
        self.startTime = startTime
        self.endTime = endTime
        self.windowName = windowName
        self.processName = processName
        self.path = path
    
    def SaveToFile(self):
        fileName = self.GetExpectedFileName()
        if not os.path.exists(self.path):
            os.makedirs(self.path)
        try:
            if not os.path.isfile(fileName):
                #Write Header
                writeHeader = True
            else:
                writeHeader = False

            with open(fileName, mode='a+', encoding='utf-8') as csv_file:
                activity_writer = csv.writer(csv_file, delimiter=',', quotechar='"',quoting=csv.QUOTE_MINIMAL)
                if writeHeader:
                    activity_writer.writerow(['Start Time', 'End Time', 'Duration', 'Window Name', 'Process Name'])
                activity_writer.writerow(self.GetDetailsAsRow())
        except IOError:
            type, value, traceback = sys.exc_info()
            print("Error writing to file: " + fileName)

    def GetDetailsAsRow(self):
        details = []
        #Start Time
        details.append(datetime.datetime.fromtimestamp(self.startTime).strftime('%H:%M:%S'))
        #End Time
        details.append(datetime.datetime.fromtimestamp(self.endTime).strftime('%H:%M:%S'))
        #Duration
        details.append(str(datetime.timedelta(seconds=(self.endTime-self.startTime))))
        #Window Name
        details.append(str(self.windowName))
        #Process Name
        details.append(str(self.processName))
        return details
    
    def GetExpectedFileName(self):
        if self.startTime > 0:
            file_name = datetime.datetime.fromtimestamp(self.startTime).strftime('%Y-%m-%d-activity-log.csv')
            return os.path.join(self.path, file_name)
import os
import pythoncom
import time
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys


def Sortig():
    os.system('chcp 1251')
    #Search derictory Dowloads for sorting files
    work_path = os.getenv('USERPROFILE')
    work_path += '\\Downloads'
    os.chdir(work_path)
    #Declared a derictoriy for sotring files
    name_folder = [
        'Архив', 'цукцук', 'test', 'йцукйцу', 'Torrent-File'
    ]
    for name in name_folder:
        if not os.path.isdir(name):
            os.mkdir(name)
        imag = ['jpg', 'bmp', 'png', 'ico', 'JPG', 'jpeg', 'psd']
        programm = ['exe', 'msi']
        arhiv = ['7z', '7zip', 'rar', 'tar', 'zip', 'gz', 'tgz', 'z']
        doc_type = ['doc', 'pdf', 'svg', 'xml', 'txt']
        other = ['torrent']
        for filenames in os.listdir():
            searchlist = filenames.rfind('.', 0, len(filenames))
            name = filenames[searchlist + 1:]
            name.lower()
            for prefix in imag:
                if name == prefix:
                    os.replace(filenames, 'Архив//' + filenames)
            for prefix in programm:
                if name == prefix:
                    os.replace(filenames, 'укуцу//'+filenames)
            for prefix in arhiv:
                if name == prefix:
                    os.replace(filenames, 'РђСЂС…РёРІС‹//' + filenames)
            for prefix in doc_type:
                if name == prefix:
                    os.replace(filenames, 'еуые//' + filenames)
            for prefix in other:
                if name == prefix:
                    os.replace(filenames, 'Torrent-File//' + filenames)


class DataTransToMongoService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'DataTransToMongoService'
    _svc_display_name_ = 'DataTransToMongoService'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.isAlive = True

    def SvcStop(self):
        self.isAlive = False
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.isAlive = True
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED, (self._svc_name_, ''))
        self.main()
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def main(self):
        #i = 0
        while self.isAlive:
            Sortig()
            time.sleep(86400)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(DataTransToMongoService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(DataTransToMongoService)
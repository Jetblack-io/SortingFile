import os
import time
import win32serviceutil
import win32service
import win32event
import servicemanager


def CreateDerictory():
    #Search derictory Dowloads for sorting files
    work_path = os.getenv('USERPROFILE')
    work_path += '\\Downloads'
    os.chdir(work_path)
    #Declared a derictoriy for sotring files
    name_folder = [
        'Архив', 'Картинки', 'Программы', 'Документы', 'Torrent-File'
    ]
    for name in name_folder:
        if not os.path.isdir(name):
            os.mkdir(name)


def Sortig():
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
                os.replace(filenames, 'Картинки//' + filenames)
        for prefix in programm:
            if name == prefix:
                os.replace(filenames, 'Программы//' + filenames)
        for prefix in arhiv:
            if name == prefix:
                os.replace(filenames, 'Архив//' + filenames)
        for prefix in doc_type:
            if name == prefix:
                os.replace(filenames, 'Документы//' + filenames)
        for prefix in other:
            if name == prefix:
                os.replace(filenames, 'Torrent-File//' + filenames)


class ServiceSorting(win32serviceutil.ServiceFramework):
    _svc_name_ = 'File-Sorting'
    _svc_display_name_ = 'Служба сортировки файлов'
    _svc_description_ = 'Python Service sorting file rof directory USER/Downloads'


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
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        self.main()

    def main(self):
        os.system('chcp 866')
        CreateDerictory()
        while self.isAlive:
#            try:
            Sortig()
#            except Exception as e:
#                pass
            time.sleep(30)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ServiceSorting)
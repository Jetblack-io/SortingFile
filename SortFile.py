import os
import time
from SMWinservice import SMWinservice


class Sorting(SMWinservice):
    _svc_name_ = "Sortig-Service"
    _svc_display_name_ = "Sortig Service"
    _svc_description_ = "Тест сообщение "

    def Start(self):
        self.isrunning = True

    def Stop(self):
        self.isrunning = True

    def main(self):
        #Search derictory Dowloads for sorting files
        work_path = os.getenv('USERPROFILE')
        work_path += '\\Downloads'
        os.chdir(work_path)
        while self.isrunning:
            #Declared a derictoriy for sotring files
            name_folder = ['Архивы', 'Картинки', 'Документы', 'Программы','Torrent-File']
            for name in name_folder:
                if not os.path.isdir(name):
                    os.mkdir(name)
            imag = ['jpg', 'bmp', 'png', 'ico','JPG','jpeg','psd']
            programm = ['exe', 'msi']
            arhiv = ['7z', '7zip', 'rar','tar','zip','gz','tgz','z']
            doc_type = ['doc', 'pdf', 'svg','xml','txt']
            other = ['torrent']
            for filenames in os.listdir():
                searchlist = filenames.rfind('.', 0, len(filenames))
                name = filenames[searchlist+1:]
                name.lower()
                for prefix in imag:
                    if name == prefix:
                        os.replace(filenames,'Картинки//'+filenames)
                for prefix in programm:
                    if name == prefix:
                        os.replace(filenames,'Программы//'+filenames)
                for prefix in  arhiv:
                    if name == prefix:
                        os.replace(filenames,'Архивы//'+filenames)
                for prefix in  doc_type:
                    if name == prefix:
                        os.replace(filenames,'Документы//'+filenames)
                for prefix in  other:
                    if name == prefix:
                        os.replace(filenames,'Torrent-File//'+filenames)
            time.sleep(30)

if __name__ == '__main__':
    Sorting.parse_command_line()
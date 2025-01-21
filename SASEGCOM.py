from time import sleep
from datetime import datetime
import win32com.client
import os
from pathlib import Path

class SASEGHandler:
    file_list = None
    PROFILE_NAME = ''
    EG_VERSION = ''

    def __init__(self, file_list=None, folder='',PROFILE_NAME="sasserver", EG_VERSION="8.1"):
        """

        :param file_list: list of filename strings ending with .sas
        :param PROFILE_NAME: default is sasserver
        :param EG_VERSION: default is 8.1
        """
        super(SASEGHandler, self).__init__()
        self.current_folder = folder
        self.file_list = file_list
        self.PROFILE_NAME = PROFILE_NAME
        self.EG_VERSION = EG_VERSION

    def batch_run(self, root=None):
        # Collect all items from List_R and perform your custom action
        print('===== SASEGCOM running: ======')
        if len(self.file_list) == 0:
            return
        for index,file in enumerate(self.file_list):
            root.event_generate("<<event1>>", when='tail', state=index)
            # Extract filename without extension
            file_name = Path(file).stem
            # cwd = os.getcwd()
            cwd = self.current_folder
            realcwd1 = cwd.replace("\'", "")
            realcwd2 = realcwd1.replace("\\", "/")
            realcwd3 = realcwd2.replace("Z:", "/data1")
            file_path = realcwd3 + '/' + file

            file_path_raw = cwd + '/' + file

            # for log: EG scripting bug, it always picks current disk
            # thus we cannot run it from local, otherwise it starts with C:/
            realcwd4 = realcwd2.replace("Z:", "")
            log_name = realcwd4 + '/' + file_name + '.log'
            log_name_full = realcwd2 +'/' + file_name + '.log'

            if os.path.isfile(file_path_raw):
                now = datetime.now()
                current_time = now.strftime("%H:%M:%S")
                print(f'{current_time} Batching: {file_path_raw}')
                app = win32com.client.Dispatch(f"SASEGObjectModel.Application.{self.EG_VERSION}")
                app.SetActiveProfile(self.PROFILE_NAME)
                project = app.New()

                # write code to the new file and run
                codeItem = project.CodeCollection.Add()
                codeItem.Server = "SASApp_UTF8"
                codeItem.Text = f"%include \"{file_path}\";"
                codeItem.Run()
                codeItem.Log.SaveAs(log_name)

                # convert log file name to utf-8
                log_name_full = realcwd2 + '/' + file_name + '.log'
                print('Transcoding from GB2312 to UTF-8-BOM')
                self.gbk_to_utf8(filename=log_name_full, newFilename=log_name_full)
                print('Batch Done')
            else:
                print('file path not recognized by os: ')
                print('skipped: ' + file_path)
                continue


    def batch_run_dummy(self, status_list=None, root=None):
        print('=====Test Batching=====')
        if len(self.file_list) == 0 or root is None:
            if len(self.file_list) == 0:
                print('no file to batch')
            else:
                print('did not get root')
            return

        for index, file in enumerate(self.file_list):
            # Extract filename without extension
            file_name = Path(file).stem
            # cwd = os.getcwd()
            cwd = self.current_folder
            realcwd1 = cwd.replace("\'", "")
            realcwd2 = realcwd1.replace("\\", "/")
            realcwd3 = realcwd2.replace("Z:", "/data1")
            file_path = realcwd3 + '/' + file
            print('batching: '+file_path)
            root.event_generate("<<event1>>", when='tail', state=index)
            sleep(2)


    # TODO: need to test on large log file. Should I read by chunk ?
    def gbk_to_utf8(self, filename='', newFilename='', encoding_from='GB2312', encoding_to='UTF-8-sig'):
        with open(filename, 'r', encoding=encoding_from) as fr:
            content = fr.read()

        with open(newFilename, 'w', encoding=encoding_to) as fw:
            fw.write(content)


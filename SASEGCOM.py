from time import sleep

import win32com.client
import os
from pathlib import Path

class SASEGHandler:
    file_list = None
    PROFILE_NAME = ''
    EG_VERSION = ''

    def __init__(self, file_list=None, PROFILE_NAME="sasserver", EG_VERSION="8.1"):
        """

        :param file_list: list of filename strings ending with .sas
        :param PROFILE_NAME: default is sasserver
        :param EG_VERSION: default is 8.1
        """
        super(SASEGHandler, self).__init__()
        self.file_list = file_list
        self.PROFILE_NAME = PROFILE_NAME
        self.EG_VERSION = EG_VERSION

    def batch_run(self):
        # Collect all items from List_R and perform your custom action
        if len(self.file_list) == 0:
            return
        for file in self.file_list:
            # Extract filename without extension
            file_name = Path(file).stem
            cwd = os.getcwd()
            realcwd1 = cwd.replace("\'", "")
            realcwd2 = realcwd1.replace("\\", "/")
            realcwd3 = realcwd2.replace("Z:", "/data1")
            file_path = realcwd3 + '/' + file
            # file_path = os.path.join(realcwd3,file)
            # dir_path = os.path.dirname(os.path.realpath(__file__))
            # print(file_path)
            # print(dir_path)
            # print(cwd)
            app = win32com.client.Dispatch(f"SASEGObjectModel.Application.{self.EG_VERSION}")
            app.SetActiveProfile(self.PROFILE_NAME)
            project = app.New()

            # write code to the new file and run
            codeItem = project.CodeCollection.Add()
            codeItem.Server = "SASApp_UTF8"
            codeItem.Text = f"%include \"{file_path}\";"
            codeItem.Run()
            codeItem.Log.SaveAs(f"{file_name}.log")

    def batch_run_dummy(self, status_list=None):
        print('=====Batching=====')
        i = 0
        for file in self.file_list:
            sleep(1)
            print(file)
            status_list.append(i)
            i + 1

        status_list = []



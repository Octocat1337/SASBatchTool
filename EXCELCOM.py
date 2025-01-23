import win32com.client
import os
from time import sleep

class EXCELHandler:
    current_folder = ''
    tracker_folder = ''
    tracker_file = ''

    # Global Vars
    sdtm_sheet_name = 'SDTM Dataset'
    tlf_sheet_name = 'TLF'

    topline_pos = 0
    intext_pos = 0
    combine_pos = 0
    program_pos = 0
    qcprogram_pos = 0

    is_qc = False

    def __init__(self, folder='', dummy=False):
        """
            :param folder: current folder path
        """
        self.current_folder = folder

        if dummy:
            return

        if '\\qc\\' in self.current_folder or '/qc/' in self.current_folder:
            self.is_qc = True

        self.tracker_folder = self.get_tracker_folder(path=self.current_folder)

        # print('tracker_folder = ' + self.tracker_folder)

        # get tracker excel file
        file_list_tmp = os.listdir(self.tracker_folder)
        # print(file_list_tmp)
        for item in file_list_tmp:
            if 'tracker.xls' in item:
                self.tracker_file = item
                break
        print('found tracker file ' + self.tracker_file)

        self.tracker_file_path = os.path.join(self.tracker_folder,self.tracker_file)
        print(self.tracker_file_path)

    def get_filelist_dummy(self, type='', root=None):
        root.event_generate("<<event2>>", when='tail', state=1)
        sleep(5)
        root.event_generate("<<event2>>", when='tail', state=2)
        list = []
        return list

    def get_filelist(self, type='', root=None):
        root.event_generate("<<event2>>", when='tail', state=1)
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        excel.ScreenUpdating = False

        file = self.tracker_file_path

        workbook = excel.Workbooks.open(file)
        # excel.Visible = False
        # workbook.Windows(1).Visible = False

        sdtm_sheet_name = workbook.Worksheets(self.sdtm_sheet_name)

        TLF_sheet = workbook.WorkSheets(self.tlf_sheet_name)

        row_end = TLF_sheet.UsedRange.Rows.Count
        col_end = TLF_sheet.UsedRange.Columns.Count

        print('rows:' + str(row_end) + ' cols:' + str(col_end))

        # note: excel representation number starts at 1 instead of 0 !

        for num in range(col_end):
            header = TLF_sheet.Rows.Item(1).Columns.Item(num + 1).Text
            if header.lower() == 'combine':
                self.combine_pos = num + 1
            elif header.lower() == 'topline':
                self.topline_pos = num + 1
            elif header.lower() == 'in-text':
                self.intext_pos = num + 1
            elif header.lower() == 'program':
                self.program_pos = num + 1
            elif header.lower() == 'qc program':
                self.qcprogram_pos = num + 1
            elif header.lower() == 'fn1':
                break

        return_list = []
        if type == 'topline':
            # get topline
            for num in range(row_end):
                row = num + 1
                istopline = TLF_sheet.Rows.Item(row).Columns.Item(self.topline_pos).Text
                if istopline.lower() == 'y':
                    if self.is_qc:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.qcprogram_pos).Text + '.sas')
                    else:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.program_pos).Text + '.sas')


        if type == 'combine':
            # get combine
            for num in range(row_end):
                row = num + 1
                iscombine = TLF_sheet.Rows.Item(row).Columns.Item(self.combine_pos).Text
                if iscombine.lower() == 'y':
                    if self.is_qc:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.qcprogram_pos).Text + '.sas')
                    else:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.program_pos).Text + '.sas')

        if type == 'in-text':
            # get combine
            for num in range(row_end):
                row = num + 1
                isin_text = TLF_sheet.Rows.Item(row).Columns.Item(self.intext_pos).Text
                if isin_text.lower() == 'y':
                    if self.is_qc:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.qcprogram_pos).Text + '.sas')
                    else:
                        return_list.append(TLF_sheet.Rows.Item(row).Columns.Item(self.program_pos).Text + '.sas')


        # in the end, close the workbook and the excel
        workbook.Close(SaveChanges=False)
        excel.Quit()
        root.event_generate("<<event2>>", when='tail', state=2)
        return return_list

    def get_tracker_folder(self, path='',target_folder='program'):
        '''
        a\b\c\d\dryrun1\program\tlf -> a\b\c\d\dryrun1
        :return: dryrun1 level folder
        '''
        # print('getting tracker folder')
        normpath = os.path.normpath(path)
        parts = normpath.split(os.sep)  # Split the path into parts using the OS separator
        # print(parts)
        try:
            index = parts.index(target_folder)
            base_path = ''
            if self.is_qc:
                # base_path = os.path.join(*parts[:index -1],'document','tracker')
                base_path = os.path.join('Z:\\',*parts[1:index-1],'document','tracker')
            else:
                # base_path = os.path.join(*parts[:index],'document','tracker')
                base_path = os.path.join('Z:\\',*parts[1:index],'document','tracker')
            return base_path
        except ValueError:
            return ''  # Target folder not found


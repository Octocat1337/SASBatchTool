import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import SASEGCOM
import json


class MainWindow:
    # Modules
    EXCELEG = None
    SASEG = None  #TODO
    #Global Vars
    current_folder = ''
    current_folder_files = []
    # Initialize the main window
    root = None

    outer_frame = None
    top_frame = None
    main_frame = None
    bot_frame = None

    left_frame = None
    mid_frame = None
    right_frame = None
    tool_frame = None

    # two lists
    left_listbox = None
    right_listbox = None
    left_list = []
    right_list = []
    # two label above the 2 lists
    left_label = None
    right_label = None

    # Move up, Move down buttons for lists
    btn_up = None
    btn_down = None
    btn_top = None
    btn_bottom = None

    # middle 2 buttons
    btn_ltr = None
    btn_rtl = None

    # bottom button
    btn_batch = None
    # bottom progress bar
    progress_bar = None

    def __init__(self):
        self.testnum = 0
        self.root = tk.Tk()
        self.root.title("Batch Tool")
        self.root.geometry("600x500")
        self.root.maxsize(800, 700)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.outer_frame = tk.Frame(self.root, padx=10, pady=10, bg="seashell3")
        self.outer_frame.columnconfigure(0, weight=4)
        self.outer_frame.columnconfigure(1, weight=0)
        self.outer_frame.columnconfigure(2, weight=4)
        self.outer_frame.columnconfigure(3, weight=0)

        self.outer_frame.rowconfigure(0, weight=1)
        self.outer_frame.rowconfigure(1, weight=8)
        self.outer_frame.rowconfigure(2, weight=1)

        # 4 frames
        self.left_listbox = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED)
        self.mid_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.right_listbox = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED)
        self.functions_frame = tk.Frame(self.outer_frame, bg="seashell3")
        # left bottom and right bottom
        self.left_bottom_frame = tk.Frame(self.outer_frame, bg="seashell3")

        # labels above the two lists
        self.left_label = tk.Label(self.outer_frame, text="Will Not Batch", width=50)
        self.right_label = tk.Label(self.outer_frame, text="Will Batch", width=50)
        # label to show current folder, left bottom corner
        self.current_folder_label = tk.Label(self.left_bottom_frame, text='Curent Folder:', justify='left')
        self.current_folder_text = tk.Label(self.left_bottom_frame, text='test folder name', justify='left')

        # Buttons
        self.btn_up = tk.Button(self.mid_frame, text="move up", command=self.move_up)
        self.btn_down = tk.Button(self.mid_frame, text="move Down", command=self.move_down)
        self.btn_top = tk.Button(self.mid_frame, text="move to Top", command=self.move_to_top)
        self.btn_bottom = tk.Button(self.mid_frame, text="move to Bottom", command=self.move_to_bottom)
        self.btn_ltr = tk.Button(self.mid_frame, text=">>", command=self.move_to_right)
        self.btn_rtl = tk.Button(self.mid_frame, text="<<", command=self.move_to_left)

        # Rightmost side: select directory, batch list, etc.
        self.btn_folder = tk.Button(self.functions_frame, text='Select Folder', command=self.select_folder)
        self.btn_load_batch_list = tk.Button(self.functions_frame, text='Load Batch List', command=self.load_batch_list)
        self.btn_save_batch_list = tk.Button(self.functions_frame, text='Save Batch List', command=self.save_batch_list)
        self.btn_get_topline_tlf = tk.Button(self.functions_frame, text='Get Topline TLF', command=self.get_topline_tlf)
        self.btn_get_intext_tlf = tk.Button(self.functions_frame, text='Get In-Text TLF', command=self.get_intext_tlf)
        self.btn_get_combine_tlf = tk.Button(self.functions_frame, text='Get Combine TLF', command=self.get_combine_tlf)
        self.btn_batch = tk.Button(self.functions_frame, text="Batch", command=self.batch_run, bg="green", fg="white")

        # ScrollBars
        self.scrollbar_l = tk.Scrollbar(self.left_listbox, orient=tk.VERTICAL, command=self.left_listbox.yview)
        self.scrollbar_r = tk.Scrollbar(self.right_listbox, orient=tk.VERTICAL, command=self.right_listbox.yview)
        self.left_listbox.configure(yscrollcommand=self.scrollbar_l.set)
        self.right_listbox.configure(yscrollcommand=self.scrollbar_r.set)
        self.scrollbar_l.pack(side="right", fill=tk.Y)
        self.scrollbar_r.pack(side="right", fill=tk.Y)

        # progress Bar
        self.progress_bar = ttk.Progressbar(self.outer_frame, orient='horizontal', mode='determinate', length=100)
        # self.progress_bar.grid(row=2, column=2)

        ########## Grids: add buttons to layout ##########
        self.outer_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.left_label.grid(row=0, column=0, sticky="sw")
        self.right_label.grid(row=0, column=2, sticky="sw")
        self.left_listbox.grid(row=1, column=0, sticky="nsew")
        self.mid_frame.grid(row=1, column=1, sticky='ns')
        self.right_listbox.grid(row=1, column=2, sticky="nsew")
        self.left_bottom_frame.grid(row=2, column=0, sticky='nsew', columnspan=3)

        # mid buttons
        self.btn_top.grid(row=1, column=0, sticky="ns", padx=10, pady=10)
        self.btn_up.grid(row=2, column=0, sticky="ns", padx=10, pady=10)
        self.btn_ltr.grid(row=5, column=0, sticky="ns", padx=10, pady=10)
        self.btn_rtl.grid(row=6, column=0, sticky="ns", padx=10, pady=10)
        self.btn_down.grid(row=8, column=0, sticky="ns", padx=10, pady=10)
        self.btn_bottom.grid(row=9, column=0, sticky="ns", padx=10, pady=10)

        # Batch Button / bottom buttons: todo
        self.current_folder_label.grid(row=0, column=0, sticky='w')
        self.current_folder_text.grid(row=1, column=0, sticky='w')

        # Rightmost functional buttons
        self.functions_frame.grid(row=1, column=3, padx=5, pady=5, sticky="ns")
        self.btn_folder.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.btn_load_batch_list.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        self.btn_save_batch_list.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        self.btn_get_topline_tlf.grid(row=4, column=0, sticky="ew", padx=10, pady=10)
        self.btn_get_intext_tlf.grid(row=5, column=0, sticky="ew", padx=10, pady=10)
        self.btn_get_combine_tlf.grid(row=6, column=0, sticky="ew", padx=10, pady=10)
        self.btn_batch.grid(row=8, column=0, sticky="nsew", padx=10, pady=15)

        ########## Other setup procedures: init functions ##########
        # get current folder
        self.select_folder()
        # setup left list
        self.build_left_list_from_folder()
        print('===== Init Done =====')

    def select_folder(self, folderpath=''):
        selected_folder = ''
        # on Start-up, no folder is selected. Use current dir
        if self.current_folder == '':
            selected_folder = os.path.abspath(os.getcwd())
        elif folderpath == '':
            # print('previous dir: ' + self.current_folder)
            selected_file = filedialog.askopenfilename(
                title='Select any file to get top level folder',
                initialdir=self.current_folder,
            )
            selected_folder = os.path.dirname(selected_file)
        else:
            self.current_folder = folderpath

        if selected_folder != '':
            self.current_folder = selected_folder

        # print(f'after select folder button: {self.current_folder}')
        self.update_current_folder_text()

        # get all files ending with .sas in the current folder
        all_files = os.listdir(self.current_folder)
        all_files.sort()
        self.current_folder_files = []
        for file in all_files:
            if file.endswith(".sas"):
                self.current_folder_files.append(file)

        self.empty_left_list()
        self.empty_right_list()
        self.build_left_list_from_folder()
        self.root.update()

    def build_left_list_from_folder(self):
        '''
        Builds left list from self.current_folder
        :return:
        '''
        self.empty_left_list()

        # sort by name, TODO: move this to sort button
        # file_list_2.sort()
        # display file list in the left list
        # clear the list first
        for file in self.current_folder_files:
            self.left_listbox.insert(tk.END, file)

    def move_to_right(self):
        """
        move list items from left list to right list
        :return:
        """
        # get selected items from left: tuple
        selected_indices = self.left_listbox.curselection()
        # append to right list
        for index in selected_indices:
            self.right_listbox.insert(tk.END, self.left_listbox.get(index))

        # remove from left list, from bottom to top
        for index in selected_indices[::-1]:
            self.left_listbox.delete(index)

    def move_to_left(self):
        """
         move list items from left list to right list
         :return:
         """
        # get selected items from left: tuple
        selected_indices = self.right_listbox.curselection()
        # append to right list
        for index in selected_indices:
            color = self.right_listbox.itemcget(index,'foreground')
            self.left_listbox.insert(tk.END, self.right_listbox.get(index))
            self.left_listbox.itemconfig(self.left_listbox.size()-1 ,foreground=color)

        # remove from left list
        for index in selected_indices[::-1]:
            self.right_listbox.delete(index)

    def move_up(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        # multiple selected, cannot move up
        if len(indices) != 1:
            return
        # already at the top, cannot move up
        if indices[0] == 0:
            return

        oldpos = indices[0]
        newpos = oldpos - 1
        list_item = curr_list.get(oldpos)
        color = curr_list.itemcget(oldpos, 'foreground')
        curr_list.delete(oldpos)

        curr_list.insert(newpos, list_item)
        if color != '':
            curr_list.itemconfig(newpos, foreground=color)

        # Keep current selection
        curr_list.select_set(newpos)

    def move_down(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()
        # multiple selction, cannot move
        if len(indices) != 1:
            return
        # already at the bottom, cannot move down
        if indices[0] == curr_list.size() - 1:
            # keep selection
            curr_list.select_set(curr_list.size() - 1)
            return

        oldpos = indices[0]
        newpos = oldpos + 1
        list_item = curr_list.get(oldpos)
        color = curr_list.itemcget(oldpos, 'foreground')
        curr_list.delete(oldpos)
        curr_list.insert(newpos, list_item)
        if color != '':
            curr_list.itemconfig(newpos, foreground=color)
        curr_list.select_set(newpos)

    def move_to_top(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        if len(indices) == 0:
            return
        # move multiple to top
        i = 0
        for index in indices:
            item = curr_list.get(index)
            color = curr_list.itemcget(index, 'foreground')
            curr_list.delete(index)
            curr_list.insert(i, item)
            if color != '':
                curr_list.itemconfig(i, foreground=color)
            curr_list.select_set(i)
            i += 1

    def move_to_bottom(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        if len(indices) == 0:
            return
        # move multiple to bottom
        i = 0
        for index in indices:
            item = curr_list.get(index - i)
            color = curr_list.itemcget(index-i, 'foreground')
            curr_list.delete(index - i)
            curr_list.insert(tk.END, item)
            if color != '':
                curr_list.itemconfig(curr_list.size()-1, foreground=color)
            curr_list.select_set(curr_list.size()-1)
            i += 1

    def sort_by_name(self, file_list):
        return

    def curr_list(self) -> tk.Listbox:
        tup1 = self.left_listbox.curselection()
        if len(tup1) == 0:
            return self.right_listbox
        else:
            return self.left_listbox

    def run(self):
        self.root.mainloop()

    def batch_run(self):
        # Get right_listbox items to batch
        right_list_items = self.right_listbox.get(0, tk.END)
        if len(right_list_items) == 0:
            return
        # make a new SAS EG each time
        self.SASEG = SASEGCOM.SASEGHandler(file_list=right_list_items,folder=self.current_folder)
        # self.t = threading.Thread(target=self.batch)
        # self.t.start()
        #
        # self.progress_bar.start()
        # self.t.join()
        # self.progress_bar.stop()
        self.batch()
        messagebox.showinfo("Info", "Batch Complete")

    def batch(self):
        # self.testlist = [1]
        # self.SASEG.batch_run_dummy(status_list=[])
        self.SASEG.batch_run()

    def empty_left_list(self):
        self.left_listbox.delete(0, tk.END)

    def empty_right_list(self):
        self.right_listbox.delete(0, tk.END)

    def load_batch_list(self):
        # read in batch list file
        filetypes = [('JSON', '*.json')]
        fileextention = '.json'
        filename = filedialog.askopenfilename(defaultextension=fileextention, filetypes=filetypes)
        if filename == '':
            return
        foldername = os.path.dirname(filename)
        self.current_folder = foldername
        self.update_current_folder_text()
        with open(filename, 'r') as f:
            file_list_r = json.load(f)

        # still, read all files in the batch list file folder
        file_list_tmp = os.listdir(self.current_folder)

        file_list_l = []
        # get all files ending with .sas in the current folder
        for file in file_list_tmp:
            if file.endswith(".sas"):
                file_list_l.append(file)

        # compare with current all files, maybe something in batch list was deleted
        # Case 1. right list has something that no longer exists
        file_not_exist = list(set(file_list_r).difference(file_list_l))
        file_not_exist_set = set(file_not_exist)
        # Case 2. after case1, remove same items from left list
        file_list_l = [x for x in file_list_l if x not in set(file_list_r)]

        # in the end, build left list
        self.empty_left_list()
        for file in file_list_l:
            self.left_listbox.insert(tk.END, file)
        # build right list
        self.build_right_list(file_list_r)
        for index, item in enumerate(file_list_r):
            if item in file_not_exist_set:
                self.right_listbox.itemconfig(index, foreground='red')

    def save_batch_list(self):
        # save file name
        filetypes = [('JSON', '*.json')]
        fileextention = '.json'
        filename = filedialog.asksaveasfilename(defaultextension=fileextention, filetypes=filetypes)
        if filename == '':
            return

        # get everything from right list
        right_list = self.right_listbox.get(0, tk.END)

        # Note: keep the order in the list !
        # save the file
        batchlist = json.dumps(
            right_list,
            indent=4
        )
        with open(filename, 'w') as f:
            f.write(batchlist)

    def get_topline_tlf(self):
        pass

    def get_intext_tlf(self):
        pass

    def get_combine_tlf(self):
        pass

    def build_right_list(self, batchlist):
        self.right_listbox.delete(0, tk.END)
        for item in batchlist:
            self.right_listbox.insert(tk.END, item)

    def update_current_folder_text(self):
        self.current_folder_text.config(text=self.current_folder)

    def delete_list_item(self):
        pass

    def reset_both_lists(self):
        self.left_listbox.delete(0, tk.END)
        self.right_listbox.delete(0, tk.END)
        self.build_left_list_from_folder()
        pass

    def sort_left_list(self):
        pass

    def search_left_list(self):
        pass

if __name__ == '__main__':
    mw = MainWindow()
    mw.run()

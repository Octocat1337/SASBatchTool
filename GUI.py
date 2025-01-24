import math
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import SASEGCOM, EXCELCOM
import json
import threading


class CreateToolTip(object):
    """
    create a tooltip for a given widget
    """

    def __init__(self, widget, text='widget info'):
        self.waittime = 100  #miliseconds
        self.wraplength = 180  #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                         background="#ffffff", relief='solid', borderwidth=1,
                         wraplength=self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()


class ProgressWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.config(highlightthickness=0, bd=0)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(padx=10, pady=(10, 0))

        self.label = tk.Label(self, text="Running Excel", bg=self.cget("background"))
        self.label.pack(padx=10, pady=(0, 10))

        self.running = False

        self.update_idletasks()
        self.center_relative_to_master()  # Call the new centering function
        # Bind the <Configure> event of the master window
        self.recenter_id = self.master.bind("<Configure>", self.recenter)

    def center_relative_to_master(self):
        master_x = self.master.winfo_rootx()
        master_y = self.master.winfo_rooty()
        master_width = self.master.winfo_width()
        master_height = self.master.winfo_height()

        window_width = self.winfo_width()
        window_height = self.winfo_height()

        x = master_x + (master_width - window_width) // 2
        y = master_y + (master_height - window_height) // 2

        self.geometry(f"+{x}+{y}")

    def start(self):
        if not self.running:
            self.running = True
            self.progress.start()

    def stop(self):
        if self.running:
            self.running = False
            self.master.unbind("<Configure>",self.recenter_id)
            self.progress.stop()
            self.destroy()

    def recenter(self, event=None):  # Recenter function
        self.update_idletasks()  # Important to update the window size
        self.center_relative_to_master()


class MainWindow:
    # Modules
    EXCELHandler = None
    SASEG = None
    #Global Vars
    current_folder = ''
    current_folder_files = []
    search_performed = False
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
    btn_stop = None
    # bottom progress bar
    progress_bar = None
    batch_thread = None
    excel_thread = None
    excel_progress_window = None

    def __init__(self):
        self.testnum = 0
        self.root = tk.Tk()
        self.root.title("Batch Tool")
        self.root.geometry("800x600")
        self.root.maxsize(1000, 700)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.outer_frame = tk.Frame(self.root, padx=10, pady=10, bg="seashell3")
        self.outer_frame.columnconfigure(0, weight=4)
        self.outer_frame.columnconfigure(1, weight=0)
        self.outer_frame.columnconfigure(2, weight=4)
        self.outer_frame.columnconfigure(3, weight=0)

        self.outer_frame.rowconfigure(0, weight=1)  # search bar
        self.outer_frame.rowconfigure(1, weight=1)  # list label
        self.outer_frame.rowconfigure(2, weight=8)  # list
        self.outer_frame.rowconfigure(3, weight=1)  # current folder
        self.outer_frame.rowconfigure(4, weight=1)  # progress bar

        # main frames

        # top searchbar frame
        self.top_searchbar_frame = tk.Frame(self.outer_frame, bg="seashell3")

        self.left_listbox = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED, font=("Arial",14))
        self.mid_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.right_listbox = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED,font=("Arial",14))
        # right functions frame
        self.functions_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.functions_frame.rowconfigure(0, weight=1)
        self.functions_frame.rowconfigure(1, weight=1)
        self.functions_frame.rowconfigure(2, weight=1)
        self.functions_frame.rowconfigure(3, weight=1)
        self.functions_frame.rowconfigure(4, weight=1)
        self.functions_frame.rowconfigure(5, weight=1)
        self.functions_frame.rowconfigure(6, weight=1)
        self.functions_frame.rowconfigure(7, weight=2)
        self.functions_frame.rowconfigure(8, weight=2)

        # top searchbar
        self.searchbar = tk.Entry(self.top_searchbar_frame, width=30)
        self.searchbar.bind('<Return>', self.search_event)
        self.searchbar.pack(side=tk.LEFT, expand=True, fill=tk.Y, padx=4)
        self.btn_search = tk.Button(self.top_searchbar_frame, text='Search', padx=4, command=self.search)
        self.btn_search.pack(side=tk.LEFT, padx=4, expand=True, fill=tk.Y)
        self.closeimg = tk.PhotoImage(file='./src/close16.png')
        self.btn_reset_searchbar = tk.Button(self.top_searchbar_frame, image=self.closeimg, command=self.reset_search)
        self.btn_reset_searchbar.pack(side=tk.LEFT, padx=4, expand=True, fill=tk.Y)
        self.btn_reset_searchbar.config(state="disabled")

        # left bottom and right bottom
        self.left_bottom_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.left_bottom_frame.rowconfigure(0, weight=1)
        self.left_bottom_frame.rowconfigure(1, weight=1)
        self.left_bottom_frame.columnconfigure(0, weight=1)
        # labels above the two lists
        self.left_label_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.right_label_frame = tk.Frame(self.outer_frame, bg="seashell3")

        self.sort_img = tk.PhotoImage(file='./src/sort-arrow-up16.png')

        self.btn_sort_left = tk.Button(
            self.left_label_frame, image=self.sort_img, command=self.sort_left_list)
        self.btn_sort_right = tk.Button(
            self.right_label_frame, image=self.sort_img, command=self.sort_right_list)

        self.left_label = tk.Label(self.left_label_frame, text="Will Not Batch", width=50)
        self.right_label = tk.Label(self.right_label_frame, text="Will Batch", width=50)

        self.btn_sort_left.pack(side=tk.LEFT, padx=0, expand=True, fill=tk.Y)
        self.left_label.pack(side=tk.LEFT, padx=(2, 0), expand=True, fill=tk.Y)

        self.btn_sort_right.pack(side=tk.LEFT, padx=0, expand=True, fill=tk.Y)
        self.right_label.pack(side=tk.LEFT, padx=(2, 0), expand=True, fill=tk.Y)

        self.tooltip_left = CreateToolTip(self.btn_sort_left, 'sort by name')
        self.tooltip_right = CreateToolTip(self.btn_sort_right, 'sort by name')

        self.left_label_frame.grid(row=1, column=0, sticky='sw')
        self.right_label_frame.grid(row=1, column=2, sticky='sw')
        # label to show current folder, left bottom corner
        self.current_folder_label = tk.Label(self.left_bottom_frame, text='Curent Folder:', justify='left')
        self.current_folder_text = tk.Label(self.left_bottom_frame, text='test folder name', justify='left')

        # Middle Frame Buttons
        # images
        # self.left_arrow_img = tk.PhotoImage(file='./src/arrows/thin-long-left-arrow32.png')

        self.btn_up = tk.Button(self.mid_frame, text="move up", command=self.move_up,width=16, height=1)
        self.btn_down = tk.Button(self.mid_frame, text="move Down", command=self.move_down,width=16, height=1)
        self.btn_top = tk.Button(self.mid_frame, text="move to Top", command=self.move_to_top,width=16, height=1)
        self.btn_bottom = tk.Button(self.mid_frame, text="move to Bottom", command=self.move_to_bottom,width=16, height=1)
        self.btn_ltr = tk.Button(self.mid_frame, text=">>", command=self.move_to_right,width=16, height=1)
        self.btn_rtl = tk.Button(self.mid_frame, text="<<", command=self.move_to_left,width=16, height=1)
        # self.btn_rtl = tk.Button(self.mid_frame, image=self.left_arrow_img, command=self.move_to_left, width=64, height=32)

        # Rightmost side: select directory, batch list, etc.
        self.btn_folder = tk.Button(self.functions_frame, text='Select Folder', command=self.select_folder)
        self.btn_load_batch_list = tk.Button(self.functions_frame, text='Load Batch List', command=self.load_batch_list)
        self.btn_save_batch_list = tk.Button(self.functions_frame, text='Save Batch List', command=self.save_batch_list)
        self.btn_get_topline_tlf = tk.Button(self.functions_frame, text='Get Topline TLF', command=self.get_topline_tlf)
        self.btn_get_intext_tlf = tk.Button(self.functions_frame, text='Get In-Text TLF', command=self.get_intext_tlf)
        self.btn_get_combine_tlf = tk.Button(self.functions_frame, text='Get Combine TLF', command=self.get_combine_tlf)
        # self.btn_batch = tk.Button(self.functions_frame, text="Batch", command=self.batch_run, bg="green", fg="white")
        self.btn_batch = tk.Button(self.functions_frame, text="Batch",
                                   command=self.run_batch_thread, bg="#385f30", fg="white", font=("Arial", 16))
        self.btn_stop = tk.Button(self.functions_frame, text="Stop",
                                  command=self.stop_batch, bg="#803328", fg="white", font=("Arial", 16))
        # ScrollBars
        self.scrollbar_l = tk.Scrollbar(self.left_listbox, orient=tk.VERTICAL, command=self.left_listbox.yview)
        self.scrollbar_r = tk.Scrollbar(self.right_listbox, orient=tk.VERTICAL, command=self.right_listbox.yview)
        self.left_listbox.configure(yscrollcommand=self.scrollbar_l.set)
        self.right_listbox.configure(yscrollcommand=self.scrollbar_r.set)
        self.scrollbar_l.pack(side="right", fill=tk.Y)
        self.scrollbar_r.pack(side="right", fill=tk.Y)

        # progress Bar
        self.progress_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.progress_label = tk.Label(self.progress_frame, text='Batching:', width=30, anchor='w',
                                       background='seashell3')
        self.progress_bar = ttk.Progressbar(self.progress_frame,
                                            orient='horizontal', mode='determinate', length=200)
        # self.progress_label.pack(side=tk.LEFT, padx=0, expand=False, fill=tk.Y, anchor='nw')
        # self.progress_bar.pack(side=tk.LEFT, padx=(20, 0), expand=True, fill=tk.Y, anchor='nw')

        self.progress_frame.rowconfigure(0, weight=1)
        self.progress_frame.rowconfigure(1, weight=1)
        self.progress_frame.columnconfigure(0, weight=1)

        self.progress_label.grid(row=0, column=0, sticky='ew')
        self.progress_bar.grid(row=1, column=0, sticky='ew')

        # SASEG event progressbar
        self.root.bind('<<event1>>', self.update_progress_bar)
        # EXCEL event progress window
        self.root.bind('<<event2>>', self.progress_window)

        ########## Grids: add buttons to layout ##########
        self.outer_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.top_searchbar_frame.grid(row=0, column=0, sticky='sw', columnspan=3)

        self.left_listbox.grid(row=2, column=0, sticky="nsew")
        self.mid_frame.grid(row=2, column=1, sticky='ns')
        self.right_listbox.grid(row=2, column=2, sticky="nsew")

        self.left_bottom_frame.grid(row=3, column=0, sticky='nsew', columnspan=3)
        self.progress_frame.grid(row=4, column=0, columnspan=3, sticky='ew')

        self.functions_frame.grid(row=2, column=3, padx=5, pady=5, sticky="ns", rowspan=3)

        # mid buttons
        self.btn_top.grid(row=1, column=0, sticky="ns", padx=10, pady=10)
        self.btn_up.grid(row=2, column=0, sticky="ns", padx=10, pady=10)
        self.btn_ltr.grid(row=5, column=0, sticky="ns", padx=10, pady=10)
        self.btn_rtl.grid(row=6, column=0, sticky="ns", padx=10, pady=10)
        self.btn_down.grid(row=8, column=0, sticky="ns", padx=10, pady=10)
        self.btn_bottom.grid(row=9, column=0, sticky="ns", padx=10, pady=10)

        # Batch Button / bottom buttons: todo
        self.current_folder_label.grid(row=0, column=0, sticky='w')
        self.current_folder_text.grid(row=1, column=0, sticky='ew')

        # Rightmost functional buttons
        self.btn_folder.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        self.btn_load_batch_list.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.btn_save_batch_list.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        self.label_placeholder = tk.Label(self.functions_frame, text='', background='seashell3')
        self.label_placeholder.grid(row=3, column=0, sticky='nsew')

        self.btn_get_topline_tlf.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
        self.btn_get_intext_tlf.grid(row=5, column=0, sticky="ew", padx=10, pady=5)
        self.btn_get_combine_tlf.grid(row=6, column=0, sticky="ew", padx=10, pady=5)
        self.btn_batch.grid(row=7, column=0, sticky="nsew", padx=10, pady=5)
        self.btn_stop.grid(row=8, column=0, sticky="nsew", padx=10, pady=5)
        ########## Other setup procedures: init functions ##########
        # get current folder
        self.select_folder()
        # setup left list
        self.build_left_list_from_folder()
        self.center_window()  # Center the window
        print('===== Init Done =====')

    def center_window(self):
        # pass
        # self.root.eval('tk::PlaceWindow %s center' % self.root.winfo_toplevel())
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()

        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"+{x}+{y}")

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
            color = self.right_listbox.itemcget(index, 'foreground')
            self.left_listbox.insert(tk.END, self.right_listbox.get(index))
            self.left_listbox.itemconfig(self.left_listbox.size() - 1, foreground=color)

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
            color = curr_list.itemcget(index - i, 'foreground')
            curr_list.delete(index - i)
            curr_list.insert(tk.END, item)
            if color != '':
                curr_list.itemconfig(curr_list.size() - 1, foreground=color)
            curr_list.select_set(curr_list.size() - 1)
            i += 1

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
        right_list_items = list(self.right_listbox.get(0, tk.END))
        if len(right_list_items) == 0:
            return
        # make a new SAS EG each time
        self.SASEG = SASEGCOM.SASEGHandler(file_list=right_list_items, folder=self.current_folder)

        self.batch()

        self.progress_label.config(text='Done')
        self.progress_bar.config(value=100)

        messagebox.showinfo("Info", "Batch Complete")

    def batch(self):
        # self.testlist = [1]
        # self.SASEG.batch_run_dummy(status_list=[],root=self.root)
        self.SASEG.batch_run(root=self.root)

    def stop_batch(self):
        if self.SASEG is None:
            return
        self.SASEG.stop = True

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
        if 'program' in self.current_folder \
                and ('tlf' in self.current_folder or 'qctlf' in self.current_folder):
            self.excel_thread = threading.Thread(target=self.get_topline_tlf_run)
            self.excel_thread.daemon = True
            self.excel_thread.start()
        else:
            return

    def get_topline_tlf_run(self):
        # self.EXCELHandler = EXCELCOM.EXCELHandler(folder=self.current_folder,dummy=True)
        self.EXCELHandler = EXCELCOM.EXCELHandler(folder=self.current_folder)
        # file_list_r = self.EXCELHandler.get_filelist(type='topline')
        file_list_r = self.EXCELHandler.get_filelist_dummy(type='topline', root=self.root)

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

    def get_intext_tlf(self):
        if 'program' in self.current_folder \
                and ('tlf' in self.current_folder or 'qctlf' in self.current_folder):
            self.excel_thread = threading.Thread(target=self.get_intext_tlf_run)
            self.excel_thread.daemon = True
            self.excel_thread.start()

    def get_intext_tlf_run(self):
        self.EXCELHandler = EXCELCOM.EXCELHandler(folder=self.current_folder)
        file_list_r = self.EXCELHandler.get_filelist(type='in-text', root=self.root)

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

    def get_combine_tlf(self):
        if 'program' in self.current_folder \
                and ('tlf' in self.current_folder or 'qctlf' in self.current_folder):
            self.excel_thread = threading.Thread(target=self.get_combine_tlf_run)
            self.excel_thread.daemon = True
            self.excel_thread.start()

    def get_combine_tlf_run(self):
        self.EXCELHandler = EXCELCOM.EXCELHandler(folder=self.current_folder)
        file_list_r = self.EXCELHandler.get_filelist(type='combine', root=self.root)

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

    def build_left_list(self, batchlist):
        self.left_listbox.delete(0, tk.END)
        for item in batchlist:
            self.left_listbox.insert(tk.END, item)

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
        left_list = list(self.left_listbox.get(0, tk.END))
        left_list.sort()
        self.empty_left_list()
        for file in left_list:
            self.left_listbox.insert(tk.END, file)

    def sort_right_list(self):
        right_list = list(self.right_listbox.get(0, tk.END))
        right_list.sort()
        self.empty_right_list()
        for file in right_list:
            self.right_listbox.insert(tk.END, file)

    def search_event(self, event):
        self.search()

    def search(self):
        '''
        searches both lists and display results
        not sure what to do
        :return:
        '''

        # 1. get search text
        text = self.searchbar.get()

        if text == '':
            # Case 1: User searched nothing
            if not self.search_performed:
                return
            # Case 2: After normal search, user wants to proceed
            else:
                self.reset_search()
                self.btn_search.config(state='normal')
                self.btn_reset_searchbar.config(state='disabled')
        else:
            self.search_performed = True
            self.btn_search.config(state='disabled')
            self.btn_reset_searchbar.config(state='normal')

        # 2. record current state
        self.left_list = list(self.left_listbox.get(0, tk.END))
        self.right_list = list(self.right_listbox.get(0, tk.END))

        # 3. search two lists, remove found results from original lists
        result_left = []
        result_right = []

        new_left_list = []
        new_right_list = []

        for item in self.left_list:
            if text in item:
                result_left.append(item)
            # else:
            #     new_left_list.append(item)

        for item in self.right_list:
            if text in item:
                result_right.append(item)
            # else:
            #     new_right_list.append(item)

        # self.left_list = new_left_list
        # self.right_list = new_right_list

        # 4. now result_ has the results, update the two list views

        self.build_left_list(result_left)
        self.build_right_list(result_right)

    def reset_search(self):
        '''
        for the reset button right to the search bar
        resets the search bar and modify the two lists
        user presses this after they've done manipulating the search results
        :return:
        '''

        text = self.searchbar.get()
        if text == '' and not self.search_performed:
            return
        else:
            self.btn_search.config(state='normal')
            self.btn_reset_searchbar.config(state='disabled')
            self.search_performed = False

        self.searchbar.delete(0, tk.END)
        left_list_tuple = self.left_listbox.get(0, tk.END)
        right_list_tuple = self.right_listbox.get(0, tk.END)

        # delete right list items from left list
        left_set = set(self.left_list)
        for item in right_list_tuple:
            if item in left_set:
                self.left_list.remove(item)

        # delete left list items from right list
        right_set = set(self.right_list)
        for item in left_list_tuple:
            if item in right_set:
                self.right_list.remove(item)

        # append to left list
        for item in left_list_tuple:
            if item not in left_set:
                self.left_list.append(item)

        # append to right list
        for item in right_list_tuple:
            if item not in right_set:
                self.right_list.append(item)

        self.build_left_list(self.left_list)
        self.build_right_list(self.right_list)

    def update_progress_bar(self, evt):
        batch_list = list(self.right_listbox.get(0, tk.END))
        total = len(batch_list)
        index = evt.state
        num = index + 1
        pct = math.ceil(num / total * 90)

        # print(evt.data)
        self.progress_label.config(text='Batching: ' + batch_list[index])
        self.progress_bar.config(value=pct)

    def run_batch_thread(self):
        self.batch_thread = threading.Thread(target=self.batch_run)
        self.batch_thread.daemon = True
        self.batch_thread.start()

    def progress_window(self, evt):
        if evt.state == 1:
            # EXCEL progress window
            self.excel_progress_window = ProgressWindow()
            self.excel_progress_window.start()
            pass
        else:
            # stop the progress bar and close the window
            self.excel_progress_window.stop()
            self.excel_progress_window = None
            pass
        pass


if __name__ == '__main__':
    mw = MainWindow()
    mw.run()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import SASEGCOM


class MainWindow:
    SASEG = None
    selected_folder = ''
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
    left_list = None
    right_list = None
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
    testnum = 0

    def __init__(self):
        self.testnum = 0
        self.root = tk.Tk()
        self.root.title("Batch Tool")
        self.root.geometry("500x500")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.outer_frame = tk.Frame(self.root, padx=10, pady=10, bg="seashell3", height=400, width=500)
        self.outer_frame.columnconfigure(0, weight=4)
        self.outer_frame.columnconfigure(1, weight=1)
        self.outer_frame.columnconfigure(2, weight=4)
        self.outer_frame.rowconfigure(0, weight=1)
        self.outer_frame.rowconfigure(1, weight=8)
        self.outer_frame.rowconfigure(2, weight=1)
        # 3 frames
        self.left_list = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED)
        self.mid_frame = tk.Frame(self.outer_frame, bg="seashell3")
        self.right_list = tk.Listbox(self.outer_frame, bg="AntiqueWhite1", selectmode=tk.EXTENDED)
        # labels above the two lists
        self.left_label = tk.Label(self.outer_frame, text="Will Not Batch")
        self.right_label = tk.Label(self.outer_frame, text="Will Batch")

        # Buttons
        self.btn_up = tk.Button(self.mid_frame, text="move up", command=self.move_up)
        self.btn_down = tk.Button(self.mid_frame, text="move Down", command=self.move_down)
        self.btn_top = tk.Button(self.mid_frame, text="move to Top", command=self.move_to_top)
        self.btn_bottom = tk.Button(self.mid_frame, text="move to Bottom", command=self.move_to_bottom)
        self.btn_ltr = tk.Button(self.mid_frame, text=">>", command=self.move_to_right)
        self.btn_rtl = tk.Button(self.mid_frame, text="<<", command=self.move_to_left)
        self.btn_batch = tk.Button(self.outer_frame, text="Batch", command=self.batch_run, bg="green", fg="white")
        # left corner select directory
        # self.btn_folder = tk.Button(self.outer_frame,text='Select Folder', command=self.select_folder)



        # ScrollBars
        self.scrollbar_l = tk.Scrollbar(self.left_list, orient=tk.VERTICAL, command=self.left_list.yview)
        self.scrollbar_r = tk.Scrollbar(self.right_list, orient=tk.VERTICAL, command=self.right_list.yview)
        self.left_list.configure(yscrollcommand=self.scrollbar_l.set)
        self.right_list.configure(yscrollcommand=self.scrollbar_r.set)
        self.scrollbar_l.pack(side="right", fill=tk.Y)
        self.scrollbar_r.pack(side="right", fill=tk.Y)

        # progress Bar
        self.progress_bar = ttk.Progressbar(self.outer_frame, orient='horizontal', mode='determinate', length=100)
        # self.progress_bar.grid(row=2, column=2)



        # Grids
        self.outer_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.left_label.grid(row=0, column=0, sticky="sw")
        self.right_label.grid(row=0, column=2, sticky="sw")
        self.left_list.grid(row=1, column=0, sticky="nsew")
        self.mid_frame.grid(row=1, column=1)
        self.right_list.grid(row=1, column=2, sticky="nsew")

        self.btn_top.grid(row=1, column=0, sticky="nsew")
        self.btn_up.grid(row=2, column=0, sticky="nsew")
        self.btn_ltr.grid(row=5, column=0, sticky="nsew")
        self.btn_rtl.grid(row=6, column=0, sticky="nsew")
        self.btn_down.grid(row=8, column=0, sticky="nsew")
        self.btn_bottom.grid(row=9, column=0, sticky="nsew")

        self.btn_batch.grid(row=2, column=1, sticky="nsew")
        # self.btn_folder.grid(row=2,column=0, sticky="nsew")
        # setup left list
        self.build_left_list()

    def select_folder(self):
        self.selected_folder = filedialog.askdirectory()
        print('printing selected folder:')
        print(self.selected_folder)
        self.empty_left_list()
        self.empty_right_list()
        self.build_left_list()
        self.root.update()


    def build_left_list(self):
        # get all files ending with .sas in the current folder
        file_list = []
        if self.selected_folder == '':
            file_list = os.listdir()
        else:
            file_list = os.listdir(self.selected_folder)

        file_list_2 = []

        for file in file_list:
            if file.endswith(".sas"):
                file_list_2.append(file)

        # sort by name
        file_list_2.sort()
        # display file list in the left list
        # clear the list first
        for file in file_list_2:
            self.left_list.insert(tk.END, file)

    def move_to_right(self):
        """
        move list items from left list to right list
        :return:
        """
        # get selected items from left: tuple
        selected_indices = self.left_list.curselection()
        # append to right list
        for index in selected_indices:
            self.right_list.insert(tk.END, self.left_list.get(index))
        # remove from left list
        for index in selected_indices[::-1]:
            self.left_list.delete(index)

    def move_to_left(self):
        """
         move list items from left list to right list
         :return:
         """
        # get selected items from left: tuple
        selected_indices = self.right_list.curselection()
        # append to right list
        for index in selected_indices:
            self.left_list.insert(tk.END, self.right_list.get(index))
        # remove from left list
        for index in selected_indices[::-1]:
            self.right_list.delete(index)

    def move_up(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        if len(indices) != 1:
            return
        # already at the top, cannot move up
        if indices[0] == 0:
            return
        oldpos = indices[0]
        newpos = oldpos - 1
        list_item = curr_list.get(oldpos)
        curr_list.delete(oldpos)

        curr_list.insert(newpos, list_item)
        # Keep current selection
        curr_list.select_set(newpos)

    def move_down(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        if len(indices) != 1:
            return
        # already at the bottom, cannot move down
        if indices[0] == curr_list.size()-1:
            curr_list.select_set(curr_list.size()-1)
            return

        oldpos = indices[0]
        newpos = oldpos + 1
        list_item = curr_list.get(oldpos)
        curr_list.delete(oldpos)
        curr_list.insert(newpos, list_item)
        curr_list.select_set(newpos)
    def move_to_top(self):
        # get selection source. which list ?
        curr_list = self.curr_list()
        indices = curr_list.curselection()

        if len(indices) == 0:
            return
        # move multiple to top
        i = 0
        j = 0
        for index in indices:
            item = curr_list.get(index)
            curr_list.delete(index)
            curr_list.insert(i, item)
            i += 1
            j += 1

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
            curr_list.delete(index - i)
            curr_list.insert(tk.END, item)
            i += 1

    def sort_by_name(self, file_list):
        return

    def curr_list(self) -> tk.Listbox:
        tup1 = self.left_list.curselection()
        if len(tup1) == 0:
            return self.right_list
        else:
            return self.left_list

    def run(self):
        self.root.mainloop()

    def batch_run(self):
        # Get right_list items to batch
        right_list_items = self.right_list.get(0, tk.END)
        if len(right_list_items) == 0:
            return
        # make a new SAS EG each time
        self.SASEG = SASEGCOM.SASEGHandler(file_list=right_list_items)
        # self.t = threading.Thread(target=self.batch)
        # self.t.start()
        #
        # self.progress_bar.start()
        # self.t.join()
        # self.progress_bar.stop()
        self.batch()
        messagebox.showinfo("Info","Batch Complete")

    def batch(self):
        # self.testlist = [1]
        # self.SASEG.batch_run_dummy(status_list=self.testlist)
        self.SASEG.batch_run()

    def empty_left_list(self):
        self.left_list.delete(0,tk.END)

    def empty_right_list(self):
        self.right_list.delete(0,tk.END)


if __name__ == '__main__':
    mw = MainWindow()
    mw.run()

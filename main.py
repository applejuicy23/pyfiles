#making comments for myself a little brainrotted

#importing moduless
import tkinter as tk
from tkinter import filedialog
import shutil
import os

selected_file = None

#choosing files
def choose_file():
    global selected_file

    file = filedialog.askopenfilename()

    if file:
        selected_file = file
        label_file.config(text=os.path.basename(file))

#moving files
def move_file():
    global selected_file

    #checking do person put file
    if not selected_file:
        label_file.config(text="Choose file first")
        return

    folder = filedialog.askdirectory()

    if folder:
        new_path = os.path.join(folder, os.path.basename(selected_file))
        shutil.move(selected_file, new_path)
        label_file.config(text="File moved")

#duplicating files
def duplicate_file():
    global selected_file

    if not selected_file:
        label_file.config(text="Choose file first")
        return

    folder = filedialog.askdirectory()

    if folder:
        new_path = os.path.join(folder, os.path.basename(selected_file))
        shutil.copy(selected_file, new_path)
        label_file.config(text="File copied")

#window
root = tk.Tk()
root.title("File Tool")
root.geometry("400x200")
root.resizable(False, False)


button_choose = tk.Button(root, text="Choose file", command=choose_file)
button_choose.grid(row=0, column=0, columnspan=2, pady=15)


label_file = tk.Label(root, text="No file selected")
label_file.grid(row=1, column=0, columnspan=2)


button_move = tk.Button(root, text="Move", command=move_file)
button_move.grid(row=2, column=0, pady=20)


button_duplicate = tk.Button(root, text="Duplicate", command=duplicate_file)
button_duplicate.grid(row=2, column=1)


root.mainloop()
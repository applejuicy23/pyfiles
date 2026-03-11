import tkinter as tk
from tkinter import filedialog
import shutil
import os

#functions
#choosing files
def choose_files():
    global selected_files
    files = filedialog.askopenfilenames()
    if files:
        selected_files = list(files)

        files_listbox.delete(0, tk.END)

        for f in selected_files:
            files_listbox.insert(tk.END, os.path.basename(f))
        status_label.config(text="Files loaded")

#choosing destination for files
def choose_destination():
    global destination_folder
    folder = filedialog.askdirectory()
    if folder:
        destination_folder = folder

        dest_listbox.delete(0, tk.END)
        dest_listbox.insert(tk.END, folder)

        status_label.config(text="Destination selected")

#moving files
def move_files():
    if not selected_files or not destination_folder:
        show_error("Select files and destination")
        return
    try:
        for f in selected_files:
            shutil.move(f, destination_folder)
        status_label.config(text="Files moved")
    except Exception as e:
        show_error(str(e))

#duplicating files
def duplicate_files():
    if not selected_files or not destination_folder:
        show_error("Select files and destination")
        return
    try:
        for f in selected_files:
            shutil.copy2(f, destination_folder)
        status_label.config(text="Files duplicated")
    except Exception as e:
        show_error(str(e))

#if error appeared
def show_error(msg):
    status_label.config(text=f"Error: {msg}", fg="red")


#window
root = tk.Tk()
root.title("PyFiles")
root.geometry("700x400")

selected_files = []
destination_folder = ""

#buttons setup
top_frame = tk.Frame(root)
top_frame.pack(pady=5)

move_btn = tk.Button(top_frame, text="Move", width=12, command=move_files)
move_btn.grid(row=0, column=0)

duplicate_btn = tk.Button(top_frame, text="Duplicate", width=12, command=duplicate_files)
duplicate_btn.grid(row=0, column=1)


#main area
main_frame = tk.Frame(root)
main_frame.pack(pady=10)


#files column
files_label = tk.Label(main_frame, text="Files")
files_label.grid(row=0, column=0)

files_listbox = tk.Listbox(main_frame, width=40, height=10)
files_listbox.grid(row=1, column=0, padx=10)

choose_files_btn = tk.Button(main_frame, text="Choose Files", command=choose_files)
choose_files_btn.grid(row=2, column=0, pady=5)


#destination column
dest_label = tk.Label(main_frame, text="To")
dest_label.grid(row=0, column=1)

dest_listbox = tk.Listbox(main_frame, width=40, height=10)
dest_listbox.grid(row=1, column=1, padx=10)

choose_dest_btn = tk.Button(main_frame, text="Choose Folder", command=choose_destination)
choose_dest_btn.grid(row=2, column=1, pady=5)

#status bar (in work)
status_label = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor="w")
status_label.pack(side=tk.BOTTOM, fill=tk.X)


root.mainloop()
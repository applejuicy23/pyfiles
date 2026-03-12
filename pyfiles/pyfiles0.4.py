import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil
import os
#var
remove_hold_job = None

selected_files = []
destination_folder = ""

#functions
#choosing files
def choose_files():
    global selected_files
    files = filedialog.askopenfilenames()
    #no-overwrite add
    if files:
        for f in files:
            if f not in selected_files:
                selected_files.append(f)
                files_listbox.insert(tk.END, os.path.basename(f))

        status_label.config(text="Files added")
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
        progress["maximum"] = len(selected_files)
        progress["value"] = 0

        for i, f in enumerate(selected_files, start=1):

            shutil.move(f, destination_folder)

            animate_progress(i)
            root.update()

        status_label.config(text="Files moved")

    except Exception as e:
        show_error(str(e))

def animate_progress(target):
    current = progress["value"]

    if current < target:
        progress["value"] += 0.2
        root.after(10, animate_progress, target)
#duplicating files
def duplicate_files():
    if not selected_files or not destination_folder:
        show_error("Select files and destination")
        return
    #animation loading
    try:
        progress["maximum"] = len(selected_files)
        progress["value"] = 0

        for i, f in enumerate(selected_files, start=1):

            shutil.copy2(f, destination_folder)

            animate_progress(i)
            root.update()

        status_label.config(text="Files duplicated")

    except Exception as e:
        show_error(str(e))
#drag&drop
def drop_files(event):
    global selected_files

    files = root.tk.splitlist(event.data)

    for f in files:
        if f not in selected_files:
            selected_files.append(f)
            files_listbox.insert(tk.END, os.path.basename(f))

    status_label.config(text="Files added via drag & drop")
#deleting files in list
def delete_selected_file():
    selection = files_listbox.curselection()
    if selection:
        index = selection[0]
        files_listbox.delete(index)
        selected_files.pop(index)
#all files clearing for files and destination list 
def clear_all_files():
    files_listbox.delete(0, tk.END)
    selected_files.clear()
    status_label.config(text="All files removed")
def clear_destination():
    global destination_folder
    destination_folder = ""
    dest_listbox.delete(0, tk.END)
#func pinch (hold click)
def start_remove_hold(event):
    global remove_hold_job
    remove_hold_job = root.after(1000, clear_all_files) 
def stop_remove_hold(event):
    global remove_hold_job
    if remove_hold_job:
        root.after_cancel(remove_hold_job)
        remove_hold_job = None

#if error appeared
def show_error(msg):
    status_label.config(text=f"Error: {msg}", fg="red")

#better to sort it but too much shit, later...
#window
root = TkinterDnD.Tk()
root.title("PyFiles")
root.geometry("600x350")
root.resizable(False, False)
style = ttk.Style()
style.theme_use("vista")

#buttons setup
top_frame = tk.Frame(root)
top_frame.pack(pady=5)

move_btn = ttk.Button(top_frame, text="Move", width=12, command=move_files)
move_btn.grid(row=0, column=0)

duplicate_btn = ttk.Button(top_frame, text="Duplicate", width=12, command=duplicate_files)
duplicate_btn.grid(row=0, column=1)


#main area
main_frame = tk.Frame(root)
main_frame.pack(pady=10)

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)


#files column
files_label = tk.Label(main_frame, text="Files")
files_label.grid(row=0, column=0)

files_listbox = tk.Listbox(main_frame, width=40, height=10)
files_listbox.grid(row=1, column=0, padx=10)
files_listbox.drop_target_register(DND_FILES)
files_listbox.dnd_bind('<<Drop>>', drop_files)

files_buttons_frame = tk.Frame(main_frame)
files_buttons_frame.grid(row=2, column=0, pady=5)

choose_files_btn = ttk.Button(files_buttons_frame, text="Choose Files", command=choose_files)
choose_files_btn.pack(side="left", padx=5)

delete_file_btn = ttk.Button(files_buttons_frame, text="Remove", command=delete_selected_file)
delete_file_btn.pack(side="left", padx=5)

delete_file_btn.bind("<ButtonPress-1>", start_remove_hold)
delete_file_btn.bind("<ButtonRelease-1>", stop_remove_hold)

#destination column
dest_label = tk.Label(main_frame, text="To")
dest_label.grid(row=0, column=1)

dest_listbox = tk.Listbox(main_frame, width=40, height=10)
dest_listbox.grid(row=1, column=1, padx=10)

dest_buttons_frame = tk.Frame(main_frame)
dest_buttons_frame.grid(row=2, column=1, pady=5)

choose_dest_btn = ttk.Button(dest_buttons_frame, text="Choose Folder", command=choose_destination)
choose_dest_btn.pack(side="left", padx=5)

delete_dest_btn = ttk.Button(dest_buttons_frame, text="Clear", command=clear_destination)
delete_dest_btn.pack(side="left", padx=5)

#status bar (in work)
bottom_frame = tk.Frame(root)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

progress = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
progress.pack(fill=tk.X, padx=5, pady=2)

status_label = tk.Label(bottom_frame, text="Ready", bd=1, relief=tk.SUNKEN, anchor="w")
status_label.pack(fill=tk.X)



root.mainloop()

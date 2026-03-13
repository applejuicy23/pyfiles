import win32gui
import win32ui
import win32con
from win32com.shell import shell, shellcon
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil
import os
import sys

#var
remove_hold_job = None
icons_cache = []
selected_files = []
destination_folder = ""

#functions
#choosing files
def choose_files():
    global selected_files

    files = filedialog.askopenfilenames()
    
    if files:
        for f in files:
            if f not in selected_files:

                selected_files.append(f)
                name = os.path.basename(f)

                icon = get_file_icon(f)

                if icon:
                    files_tree.insert("", "end", text=name, image=icon)
                    icons_cache.append(icon)
                else:
                    files_tree.insert("", "end", text=name)

                set_status("Files added")

def get_file_icon(path):

    flags = shellcon.SHGFI_ICON | shellcon.SHGFI_SMALLICON | shellcon.SHGFI_USEFILEATTRIBUTES

    ret, info = shell.SHGetFileInfo(path, 0, flags)

    hicon = info[0]

    if not hicon:
        return None

    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, 16, 16)

    hdc_mem = hdc.CreateCompatibleDC()
    hdc_mem.SelectObject(hbmp)

    win32gui.DrawIconEx(
        hdc_mem.GetHandleOutput(),
        0,
        0,
        hicon,
        16,
        16,
        0,
        None,
        win32con.DI_NORMAL
    )

    bmpinfo = hbmp.GetInfo()
    bmpstr = hbmp.GetBitmapBits(True)

    img = Image.frombuffer(
        "RGBA",
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr,
        "raw",
        "BGRA",
        0,
        1
    )

    win32gui.DestroyIcon(hicon)

    return ImageTk.PhotoImage(img) 

def open_selected_file(event):

    item = files_tree.selection()

    if not item:
        return

    item = item[0]
    index = files_tree.index(item)

    path = selected_files[index]

    try:
        os.startfile(path)
    except Exception as e:
        show_error(str(e))
#choosing destination for files
def choose_destination():
    global destination_folder

    folder = filedialog.askdirectory()

    if folder:
        folder = os.path.normpath(folder)

        destination_folder = folder

        for item in dest_tree.get_children():
            dest_tree.delete(item)

        icon = get_file_icon(folder)

        if icon:
            dest_tree.insert("", "end", text=folder, image=icon)
            icons_cache.append(icon)
        else:
            dest_tree.insert("", "end", text=folder)

        set_status("Destination selected")
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

        set_success("Files moved")

    except Exception as e:
        show_error(str(e))
#animation
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

        set_success("Files duplicated")

    except Exception as e:
        show_error(str(e))
#drag&drop
def drop_files(event):
    global selected_files

    files = root.tk.splitlist(event.data)

    for f in files:
        if f not in selected_files:
            selected_files.append(f)
            icon = get_file_icon(f)

            if icon:
                files_tree.insert("", "end", text=os.path.basename(f), image=icon)
                icons_cache.append(icon)
            else:
                files_tree.insert("", "end", text=os.path.basename(f))
    set_status("Files added via drag & drop")
#deleting files in list
def delete_selected_file():
    selection = files_tree.selection()

    if not selection:
        return

    for item in selection:
        file_name = files_tree.item(item, "text")

        index = files_tree.index(item)

        files_tree.delete(item)
        selected_files.pop(index)

        set_status(f'File "{file_name}" deleted from the list')

    if not selected_files:
        root.after(1500, lambda: set_status("Ready"))
#all files clearing for files and destination list 
def clear_all_files():
    for item in files_tree.get_children():
        files_tree.delete(item)

    selected_files.clear()

    set_status("All files removed")

    root.after(3000, lambda: set_status("Ready"))
def clear_destination():
    global destination_folder
    destination_folder = ""
    for item in dest_tree.get_children():
        dest_tree.delete(item)
#func pinch (hold click)
def start_remove_hold(event):
    global remove_hold_job
    remove_hold_job = root.after(500, clear_all_files) 
def stop_remove_hold(event):
    global remove_hold_job
    if remove_hold_job:
        root.after_cancel(remove_hold_job)
        remove_hold_job = None
#icon fixup
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def highlight_destination(event):

    item = dest_tree.identify_row(event.y)

    for i in dest_tree.get_children():
        dest_tree.item(i, tags=())

    if item:
        dest_tree.item(item, tags=("highlight",))

def dest_cursor_hover(event):

    region = dest_tree.identify("region", event.x, event.y)

    if region == "tree":
        dest_tree.config(cursor="hand2")
    else:
        dest_tree.config(cursor="")

def file_cursor_hover(event):

    region = files_tree.identify("region", event.x, event.y)

    if region == "tree":
        files_tree.config(cursor="hand2")
    else:
        files_tree.config(cursor="")

def show_file_menu(event):

    item = files_tree.identify_row(event.y)

    if item:
        files_tree.selection_set(item)
        file_menu.post(event.x_root, event.y_root)

def clear_highlight(event):

    for i in dest_tree.get_children():
        dest_tree.item(i, tags=())

def highlight_file(event):

    region = files_tree.identify("region", event.x, event.y)
    item = files_tree.identify_row(event.y)

    for i in files_tree.get_children():
        files_tree.item(i, tags=())

    if item and region == "tree":
        files_tree.item(item, tags=("highlight",))
        files_tree.config(cursor="hand2")
    else:
        files_tree.config(cursor="")

def clear_file_highlight(event):

    for i in files_tree.get_children():
        files_tree.item(i, tags=())

def delete_file_key(event):
    delete_selected_file()
def delete_destination_key(event):
    clear_destination()


def show_about():

    messagebox.showinfo(
        "About PyFiles",
        "PyFiles\n\nSimple GUI tool for file operations\nPython 3.14 + tkinter\nv0.5"
        " - UNDER CONSTRUCTION\n"
        "id/applejuicy23"
    )


def show_whats_new():

    messagebox.showinfo(
        "What's New?",
        "PyFiles v0.5\n\n"
        ">> File icons\n"
        ">> Scrollbar\n"
        ">> Double-click open\n"
        ">> 'About' and 'What's new' added\n"
    )
#if error appeared
def show_error(msg):
    status_label.config(text=f"Error: {msg}", fg="red")

def set_status(msg):
    status_label.config(text=msg, fg="black")

def set_success(msg):
    status_label.config(text=msg, fg="darkgreen")


#better to sort it but too much shit, later...
#window
root = TkinterDnD.Tk()
root.title("PyFiles")
root.geometry("600x400")
root.resizable(False, False)
root.iconbitmap(resource_path("pyfiles2.ico"))
style = ttk.Style()
style.theme_use("vista")

file_menu = tk.Menu(root, tearoff=0)

file_menu.add_command(label="Open", command=lambda: open_selected_file(None))
file_menu.add_command(label="Remove", command=delete_selected_file)
file_menu.add_separator()
file_menu.add_command(label="Reveal in Explorer", command=lambda: os.startfile(os.path.dirname(selected_files[files_tree.index(files_tree.selection()[0])])))
file_menu.add_command(label="Copy path", command=lambda: root.clipboard_append(selected_files[files_tree.index(files_tree.selection()[0])]))

#buttons setup
top_frame = tk.Frame(root)
top_frame.pack(pady=5)


move_btn = ttk.Button(top_frame, text="Move", width=12, command=move_files)
move_btn.grid(row=0, column=0)

duplicate_btn = ttk.Button(top_frame, text="Duplicate", width=12, command=duplicate_files)
duplicate_btn.grid(row=0, column=1)

about_btn = ttk.Button(top_frame, text="About", width=10, command=show_about)
about_btn.grid(row=0, column=2, padx=5)

whatsnew_btn = ttk.Button(top_frame, text="What's New", width=12, command=show_whats_new)
whatsnew_btn.grid(row=0, column=3)

#main area
main_frame = tk.Frame(root)
main_frame.pack(pady=10)

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)


#files column
files_label = tk.Label(main_frame, text="Files")
files_label.grid(row=0, column=0)

files_frame = tk.Frame(main_frame)
files_frame.grid(row=1, column=0, padx=10)

files_scrollbar = tk.Scrollbar(files_frame)
files_scrollbar.pack(side="right", fill="y")

files_tree = ttk.Treeview(
    files_frame,
    show="tree",
    height=10,
    yscrollcommand=files_scrollbar.set
)

files_tree.column("#0", width=220)
files_tree.heading("#0", text="")
files_tree.pack(side="left", fill="both", expand=True)
files_scrollbar.config(command=files_tree.yview)
files_tree.bind("<Motion>", highlight_file)
files_tree.bind("<Leave>", clear_file_highlight)
files_tree.tag_configure("highlight", background="#cce8ff")
files_tree.drop_target_register(DND_FILES)
files_tree.dnd_bind('<<Drop>>', drop_files)
files_tree.bind("<Delete>", delete_file_key)
files_tree.bind("<Button-3>", show_file_menu)

files_tree.bind("<Double-1>", open_selected_file)

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

dest_frame = tk.Frame(main_frame)
dest_frame.grid(row=1, column=1, padx=10)

dest_scrollbar = tk.Scrollbar(dest_frame)
dest_scrollbar.pack(side="right", fill="y")

dest_tree = ttk.Treeview(
    dest_frame,
    show="tree",
    height=10,
    yscrollcommand=dest_scrollbar.set
)

dest_tree.column("#0", width=220)
dest_tree.heading("#0", text="")
dest_tree.tag_configure("highlight", background="#cce8ff")
dest_tree.bind("<Motion>", dest_cursor_hover)
dest_tree.bind("<Leave>", lambda e: dest_tree.config(cursor=""))
dest_tree.bind("<Delete>", delete_destination_key)
dest_menu = tk.Menu(root, tearoff=0)

dest_menu.add_command(label="Open folder", command=lambda: os.startfile(destination_folder))
dest_menu.add_command(label="Clear", command=clear_destination)
dest_menu.add_command(label="Copy path", command=lambda: root.clipboard_append(destination_folder))

dest_tree.pack(side="left", fill="both", expand=True)
dest_scrollbar.config(command=dest_tree.yview)

files_tree.bind("<MouseWheel>", lambda e: files_tree.yview_scroll(int(-1*(e.delta/120)), "units"))
dest_tree.bind("<MouseWheel>", lambda e: dest_tree.yview_scroll(int(-1*(e.delta/120)), "units"))

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

status_label = tk.Label(
    bottom_frame,
    text="Ready",
    fg="black",
    bd=1,
    relief=tk.SUNKEN,
    anchor="w"
)
status_label.pack(fill=tk.X)

root.mainloop()
import win32gui
import win32ui
import win32con
from win32com.shell import shell, shellcon
from PIL import Image, ImageTk
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil
import os
import sys
from datetime import datetime

#var
remove_hold_job = None
icons_cache = []
selected_files = []
selected_files_set = set()
destination_folder = ""
drag_select_start = None
current_mode = "MOVE"
active_button = None
delete_hold_time = 0
delete_files_list = []

#up to release NEED TO MANAGE CODE, its hard to read and ducking my brain
#functions
#choosing files
def choose_files():
    global selected_files

    files = filedialog.askopenfilenames()
    
    if files:
        for f in files:
            if f not in selected_files_set:
                selected_files.append(f)
                selected_files_set.add(f)

                name = os.path.basename(f)

                icon = get_file_icon(f)

                if icon:
                    files_tree.insert("", "end", text=name, image=icon)
                    adjust_tree_column(files_tree, name)
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

    files_tree.delete(item)
    selected_files.pop(index)
    selected_files_set.discard(path)

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
            adjust_tree_column(dest_tree, folder)
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
        clear_all_files()

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
        clear_all_files()

    except Exception as e:
        show_error(str(e))

def set_mode(mode):
    for frame in (center_frame, create_frame, delete_frame):
        frame.pack_forget()
    global current_mode
    global active_button

    current_mode = mode

    move_btn.config(text="  MOVE")
    duplicate_btn.config(text="  COPY")
    create_btn.config(text="  CREATE")
    delete_btn.config(text="  DELETE")

    create_frame.pack_forget()
    delete_frame.pack_forget()

    for btn in (move_btn, duplicate_btn, create_btn, delete_btn):
        btn.config(relief="flat", bg="SystemButtonFace")

    if mode == "MOVE":
        move_btn.config(relief="flat", bg="#e8e8e8", text="✔ MOVE")
        active_button = move_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "DUPLICATE":
        duplicate_btn.config(relief="flat", bg="#e8e8e8", text="✔ COPY")
        active_button = duplicate_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "CREATE":
        create_btn.config(relief="flat", bg="#e8e8e8", text="✔ CREATE")
        active_button = create_btn
        create_frame.pack(fill="both", expand=True)

    elif mode == "DELETE":
        delete_btn.config(relief="flat", bg="#e8e8e8", text="✔ DELETE")
        active_button = delete_btn
        delete_frame.pack(fill="both", expand=True)

    action_btn.config(text=mode)
    

def execute_action():

    if current_mode == "MOVE":
        move_files()

    elif current_mode == "DUPLICATE":
        duplicate_files()

    elif current_mode == "DELETE":
        delete_found_files()

    elif current_mode == "CREATE":
        create_files()

#drag&drop
def drop_files(event):
    global selected_files

    files = root.tk.splitlist(event.data)

    for f in files:
        if f not in selected_files_set:

            selected_files.append(f)
            selected_files_set.add(f)

            icon = get_file_icon(f)

            if icon:
                files_tree.insert("", "end", text=os.path.basename(f), image=icon)
                icons_cache.append(icon)
                adjust_tree_column(files_tree, os.path.basename(f))
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

        path = selected_files[index]

        files_tree.delete(item)
        selected_files.pop(index)
        selected_files_set.discard(path)

        set_status(f'File "{file_name}" deleted from the list')

    if not selected_files:
        root.after(1500, lambda: set_status("Ready"))
#all files clearing for files and destination list 
def clear_all_files():
    for item in files_tree.get_children():
        files_tree.delete(item)

    selected_files.clear()
    selected_files_set.clear()

    files_tree.column("#0", width=220)

    set_status("All files removed")
    root.after(3000, lambda: set_status("Ready"))
def clear_destination():
    global destination_folder
    destination_folder = ""

    for item in dest_tree.get_children():
        dest_tree.delete(item)

    dest_tree.column("#0", width=220)
#func pinch (hold click)
def start_remove_hold(event):
    global remove_hold_job
    remove_hold_job = root.after(500, clear_all_files) 
def stop_remove_hold(event):
    global remove_hold_job
    if remove_hold_job:
        root.after_cancel(remove_hold_job)
        remove_hold_job = None

def start_delete_hold(event):
    global remove_hold_job

    remove_hold_job = root.after(
        500,
        lambda: delete_found_files_custom(delete_tree.get_children())
    )


def stop_delete_hold(event):
    global remove_hold_job

    if remove_hold_job:
        root.after_cancel(remove_hold_job)
        remove_hold_job = None


def start_drag_select(event):
    global drag_select_start
    drag_select_start = files_tree.identify_row(event.y)
def drag_select(event):

    if not drag_select_start:
        return

    current = files_tree.identify_row(event.y)

    if not current:
        return

    start_index = files_tree.index(drag_select_start)
    current_index = files_tree.index(current)

    files_tree.selection_remove(files_tree.selection())

    low = min(start_index, current_index)
    high = max(start_index, current_index)

    items = files_tree.get_children()

    for i in range(low, high + 1):
        files_tree.selection_add(items[i])
def stop_drag_select(event):
    global drag_select_start
    drag_select_start = None
#icon fixup
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def adjust_tree_column(tree, text):

    font = tkfont.nametofont("TkDefaultFont")
    text_width = font.measure(text) + 40

    if text_width > 220:
        tree.column("#0", width=text_width)
    else:
        tree.column("#0", width=220)

def adjust_delete_column_full():
    font = tkfont.nametofont("TkDefaultFont")

    max_width = 220

    for item in delete_tree.get_children():
        text = delete_tree.item(item, "text")
        width = font.measure(text) + 40

        if width > max_width:
            max_width = width

    MAX_LIMIT = 600  # чтобы не ломать UI
    delete_tree.column("#0", width=min(max_width, MAX_LIMIT))

def update_delete_scroll():
    if not delete_tree.get_children():
        delete_scroll_x.grid_remove()
        delete_container.grid_rowconfigure(1, minsize=0)
        return

    tree_width = delete_tree.winfo_width()

    if tree_width < 50:
        root.after(50, update_delete_scroll)
        return

    col_width = delete_tree.column("#0", "width")

    if col_width <= tree_width:
        delete_scroll_x.grid_remove()
        delete_container.grid_rowconfigure(1, minsize=0)
    else:
        delete_scroll_x.grid()
        delete_container.grid_rowconfigure(1, minsize=15)


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

def select_all_files(event=None):

    files_tree.selection_set(files_tree.get_children())

    return "break"

def switch_create_mode(mode):

    preview_mode_btn.config(relief="raised")
    info_mode_btn.config(relief="raised")

    if mode == "preview":
        preview_mode_btn.config(relief="sunken")

        update_preview()

        content_box.grid_remove()
        content_scroll.grid_remove()

        preview_frame.grid(row=0, column=0, sticky="nsew")

    elif mode == "info":
        info_mode_btn.config(relief="sunken")

        preview_frame.grid_remove()

        content_box.grid(row=0, column=0, sticky="nsew")
        content_scroll.grid(row=0, column=1, sticky="ns")


def show_about():

    messagebox.showinfo(
        "About PyFiles",
        "PyFiles v0.8.3t\n\nSimple GUI tool for file operations\nPython 3.14 + tkinter\nv0.8.3t"
        " - UNDER CONSTRUCTION\n"
        "id/applejuicy23"
    )


def show_whats_new():

    messagebox.showinfo(
        "What's New?",
        "PyFiles v0.8.3t\n\n"
        ">> new-UI updated\n"
        ">> Delete mode UI updated\n"
        ">> bugs fixed\n"
    )

def apply_prefix(pattern, name, num, ext):

    now = datetime.now()

    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H-%M-%S")

    result = pattern

    result = result.replace("{name}", name)
    result = result.replace("{file}", name)
    result = result.replace("{num}", str(num))
    result = result.replace("{ext}", ext)
    result = result.replace("{date}", date)
    result = result.replace("{time}", time)

    return result


def delete_found_files_custom(items):
    folder = delete_dest_entry.get().strip()

    if not folder or not os.path.isdir(folder):
        show_error("Invalid folder")
        return

    for item in items:
        name = delete_tree.item(item, "text")
        path = os.path.join(folder, name)

        try:
            os.remove(path)
            delete_tree.delete(item)
        except Exception as e:
            show_error(str(e))
            return

    set_success("Files deleted")

def create_files():

    name = file_name_entry.get()
    ext = format_entry.get()
    prefix = prefix_entry.get()
    folder = dest_entry.get()

    if not ext.startswith("."):
        ext = "." + ext

    if not folder:
        show_error("Enter destination folder")
        return

    if not os.path.isdir(folder):
        show_error("Destination folder not found")
        return

    content = content_box.get("1.0", "end")

    count = get_entry_value(count_entry, "10")

    try:
        count = int(count)
    except:
        count = 1

    count = max(1, min(count, 100))

    for i in range(count):

        new_name = apply_prefix(prefix, name, i, ext)

        filename = f"{new_name}{ext}"

        path = os.path.join(folder, filename)

        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

    set_success("Files created")
    update_preview()

def show_create_preview():

    preview = tk.Toplevel(root)
    preview.title("Preview new files")
    preview.geometry("450x350")

    preview_tree = ttk.Treeview(preview)
    preview_tree.pack(fill="both", expand=True)

    preview_tree.tag_configure("new", foreground="blue")

    name = file_name_entry.get()
    ext = format_entry.get()
    prefix = prefix_entry.get()

    count = 10

    for i in range(count):

        new_name = apply_prefix(prefix, name, i, ext)
        filename = f"{new_name}{ext}"

        preview_tree.insert("", "end", text=filename, tags=("new",))


def show_create():
    main_frame.pack(fill="both", expand=True)
    create_frame.pack(fill="both", expand=True)

def search_files(event=None):

    folder = delete_dest_entry.get().strip()
    keyword = delete_keyword_entry.get().strip()
    ext = delete_ext_entry.get().strip()

    for item in delete_tree.get_children():
        delete_tree.delete(item)

    if not folder or not os.path.isdir(folder):
        return

    try:
        files = os.listdir(folder)
    except:
        return

    for file in files:
        if keyword and keyword != "*":
            if keyword.lower() not in file.lower():
                continue

        if ext:
            if not file.endswith(ext):
                continue

        delete_tree.insert("", "end", text=file, tags=("delete",))
        adjust_delete_column_full()
        update_delete_scroll()

def delete_found_files():

    items = delete_tree.selection()

    folder = delete_dest_entry.get().strip()

    if not folder:
        show_error("Enter destination folder")
        return

    if not os.path.isdir(folder):
        show_error("Destination folder not found")
        return

    for item in items:

        name = delete_tree.item(item, "text")

        path = os.path.join(folder, name)

        try:
            os.remove(path)
            delete_tree.delete(item)

        except Exception as e:
            show_error(str(e))
            return

    set_success("Files deleted")

def add_placeholder(entry, placeholder):
    entry.insert(0, placeholder)
    entry.config(fg="gray")

    def on_focus_in(event):
        if entry.get() == placeholder and entry.cget("fg") == "gray":
            entry.delete(0, tk.END)
            entry.config(fg="black")

    def on_focus_out(event):
        if not entry.get():
            entry.insert(0, placeholder)
            entry.config(fg="gray")

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

def update_preview():

    preview_tree.delete(*preview_tree.get_children())

    name = file_name_entry.get()
    ext = format_entry.get()
    prefix = prefix_entry.get()

    count = get_entry_value(count_entry, "10")

    try:
        count = int(count)
    except:
        count = 10

    count = max(1, min(count, 100))

    font = tkfont.nametofont("TkDefaultFont")
    max_width = 220

    for i in range(count):
        new_name = apply_prefix(prefix, name, i, ext)
        filename = f"{new_name}{ext}"

        preview_tree.insert("", "end", text=filename, tags=("new",))

        width = font.measure(filename) + 40
        if width > max_width:
            max_width = width

    preview_tree.column("#0", width=max_width, stretch=False)
    preview_tree.update_idletasks()

def get_entry_value(entry, placeholder=""):
    value = entry.get().strip()
    if value == placeholder and entry.cget("fg") == "gray":
        return ""
    return value


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
root.title("PyFiles v0.8.3t")
root.geometry("700x500")
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
menu_frame = tk.Frame(root)
menu_frame.pack(fill="x")

title = tk.Label(menu_frame, text="PyFiles v0.8.3t                                                                                              ", font=("Segoe UI", 10))
title.pack(side="left", padx=10)

tabs_frame = tk.Frame(menu_frame)
tabs_frame.pack(side="left", padx=10)

def select_tab(tab):
    about_tab.config(bg="SystemButtonFace")
    whatsnew_tab.config(bg="SystemButtonFace")

    tab.config(bg="white")

about_tab = tk.Label(
    tabs_frame,
    text="ABOUT",
    padx=12,
    pady=3,
    bd=1,
    relief="raised",
    cursor="hand2"
)

about_tab.pack(side="left")

whatsnew_tab = tk.Label(
    tabs_frame,
    text="WHAT'S NEW",
    padx=12,
    pady=3,
    bd=1,
    relief="raised",
    cursor="hand2"
)

whatsnew_tab.pack(side="left", padx=(4,0))

about_tab.bind("<Button-1>", lambda e: (select_tab(about_tab), show_about()))
whatsnew_tab.bind("<Button-1>", lambda e: (select_tab(whatsnew_tab), show_whats_new()))




#main area
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

left_panel = tk.Frame(main_frame, bd=1, relief="solid")
left_panel.pack(side="left", fill="both", expand=False)

content_frame = tk.Frame(main_frame)
content_frame.pack(side="left", fill="both", expand=True)

stack_frame = tk.Frame(content_frame)
stack_frame.pack(fill="both", expand=True)

center_frame = tk.Frame(stack_frame)
center_frame.grid_columnconfigure(0, weight=1)
center_frame.grid_columnconfigure(1, weight=1)
center_frame.grid_rowconfigure(1, weight=1)


btn_style = {
    "width":14,
    "height":3,
    "font":("Segoe UI", 11),
    "bd":0,
    "anchor":"w",
    "padx":10
}

move_btn = tk.Button(left_panel, text="MOVE", command=lambda: set_mode("MOVE"), **btn_style)
move_btn.pack(fill="x", padx=0, pady=0)

duplicate_btn = tk.Button(left_panel, text="COPY", command=lambda: set_mode("DUPLICATE"), **btn_style)
duplicate_btn.pack(fill="x", padx=0, pady=0)

create_btn = tk.Button(left_panel, text="CREATE", command=lambda: set_mode("CREATE"), **btn_style)
create_btn.pack(fill="x", padx=0, pady=0)

delete_btn = tk.Button(left_panel, text="DELETE", command=lambda: set_mode("DELETE"), **btn_style)
delete_btn.pack(fill="x", padx=0, pady=0)

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

def sidebar_hover(event):
    if event.widget != active_button:
        event.widget.config(bg="#f0f0f0")

def sidebar_leave(event):
    if event.widget != active_button:
        event.widget.config(bg="SystemButtonFace")

for btn in (move_btn, duplicate_btn, create_btn, delete_btn):
    btn.bind("<Enter>", sidebar_hover)
    btn.bind("<Leave>", sidebar_leave)


#files column


files_label = tk.Label(center_frame, text="Source Files")
files_label.grid(row=0, column=0, pady=(10,5))

files_frame = tk.Frame(center_frame)
files_frame.grid(row=1, column=0, padx=20, sticky="nsew")

files_tree = ttk.Treeview(
    files_frame,
    show="tree",
    height=10,
    selectmode="extended",
)

files_tree["show"] = "tree"

files_scrollbar = tk.Scrollbar(files_frame, orient="vertical", command=files_tree.yview)
files_scrollbar_x = tk.Scrollbar(files_frame, orient="horizontal", command=files_tree.xview)

files_tree.configure(
    yscrollcommand=files_scrollbar.set,
    xscrollcommand=files_scrollbar_x.set
)

files_tree.grid(row=0, column=0, sticky="nsew")
files_scrollbar.grid(row=0, column=1, sticky="ns")
files_scrollbar_x.grid(row=1, column=0, sticky="ew")

files_frame.grid_rowconfigure(0, weight=1)
files_frame.grid_columnconfigure(0, weight=1)

files_tree.column("#0", width=220, minwidth=220, stretch=False)
files_tree.heading("#0", text="")


files_tree.bind("<Double-1>", open_selected_file)
files_tree.bind("<Motion>", highlight_file)
files_tree.bind("<Leave>", clear_file_highlight)
files_tree.bind("<Delete>", delete_file_key)
files_tree.bind("<Button-3>", show_file_menu)
files_tree.bind("<Control-a>", select_all_files)
files_tree.bind("<Button-1>", lambda e: files_tree.focus_set(), add="+")
files_tree.bind("<Button-1>", start_drag_select)
files_tree.bind("<B1-Motion>", drag_select)
files_tree.bind("<ButtonRelease-1>", stop_drag_select)
files_tree.tag_configure("highlight", background="#cce8ff")
files_tree.drop_target_register(DND_FILES)
files_tree.dnd_bind('<<Drop>>', drop_files)
files_tree.focus_set()


files_buttons_frame = tk.Frame(center_frame)
files_buttons_frame.grid_columnconfigure(0, weight=1)
files_buttons_frame.grid_columnconfigure(1, weight=1)
files_buttons_frame.grid(row=2, column=0, pady=(10,15))

choose_files_btn = ttk.Button(files_buttons_frame, text="Choose Files", command=choose_files)
choose_files_btn.grid(row=0, column=0, padx=5, sticky="e")

delete_file_btn = ttk.Button(files_buttons_frame, text="Remove", command=delete_selected_file)
delete_file_btn.grid(row=0, column=1, padx=5, sticky="w")

delete_file_btn.bind("<ButtonPress-1>", start_remove_hold)
delete_file_btn.bind("<ButtonRelease-1>", stop_remove_hold)


#destination column
dest_label = tk.Label(center_frame, text="Destination")
dest_label.grid(row=0, column=1, pady=(10,5))

dest_frame = tk.Frame(center_frame)
dest_frame.grid(row=1, column=1, padx=20, sticky="nsew")
# tree
dest_tree = ttk.Treeview(
    dest_frame,
    show="tree",
    height=10
)

files_tree["selectmode"] = "extended"

#create GUI
create_frame = tk.Frame(stack_frame)

left_create = tk.Frame(create_frame)
left_create.pack(side="left", padx=40, pady=20, anchor="n")

tk.Label(left_create, text="File name:").pack(anchor="w")

file_name_entry = tk.Entry(left_create, width=20)
file_name_entry.pack(anchor="w", pady=5)
add_placeholder(file_name_entry, "mertz")

tk.Label(left_create, text="File format:").pack(anchor="w")

format_entry = tk.Entry(left_create, width=10)
format_entry.pack(anchor="w", pady=5)
add_placeholder(format_entry, ".py")

tk.Label(left_create, text="Prefix to name").pack(anchor="w")

prefix_entry = tk.Entry(left_create, width=25)
prefix_entry.pack(anchor="w", pady=5)
add_placeholder(prefix_entry, "~{date}-1{file}cool~{num}")

tk.Label(left_create, text="Count:").pack(anchor="w")

count_entry = tk.Entry(left_create, width=10)
count_entry.pack(anchor="w", pady=5)
add_placeholder(count_entry, "10")

tk.Label(left_create, text="Destination:").pack(anchor="w")

dest_row = tk.Frame(left_create)
dest_row.pack(anchor="w", pady=5)

dest_entry = tk.Entry(dest_row, width=20)
dest_entry.pack(side="left")

def choose_create_folder():
    folder = filedialog.askdirectory()
    if folder:
        dest_entry.delete(0, tk.END)
        dest_entry.insert(0, folder)

folder_btn = tk.Button(
    dest_row,
    text="<<",
    width=4,
    bg="#ffffff",
    relief="raised",
    cursor="hand2",
    command=choose_create_folder
)

folder_btn.pack(side="left", padx=(5,0))


add_placeholder(dest_entry, "Destination folder path")

file_name_entry.bind("<KeyRelease>", lambda e: update_preview())
format_entry.bind("<KeyRelease>", lambda e: update_preview())
prefix_entry.bind("<KeyRelease>", lambda e: update_preview())
count_entry.bind("<KeyRelease>", lambda e: update_preview())



create_btn_exec = ttk.Button(
    left_create,
    text="Create",
    command=create_files
)

create_btn_exec.pack(pady=15)

right_create = tk.Frame(create_frame)
right_create.pack(side="left", padx=40, pady=20)

tk.Label(
    right_create,
    text="Contain information inside it:"
).pack(anchor="w")

mode_frame = tk.Frame(right_create)
mode_frame.pack(anchor="w", pady=(0,5))

content_container = tk.Frame(right_create)
content_container.pack(fill="both", expand=True)
content_container.grid_rowconfigure(0, weight=1)
content_container.grid_columnconfigure(0, weight=1)

preview_frame = tk.Frame(content_container)

right_create.pack(side="left", padx=40, pady=20, fill="both", expand=True)
right_create.config(width=350, height=300)

content_box = tk.Text(content_container, width=40, height=12)
content_box.config(
    font=("Consolas", 10),
    insertbackground="black",
    undo=True
)
content_box.config(tabs=("1c"))

content_scroll = tk.Scrollbar(content_container, orient="vertical", command=content_box.yview)
content_box.configure(yscrollcommand=content_scroll.set)

preview_frame.grid(row=0, column=0, sticky="nsew")
content_box.grid_remove()
content_scroll.grid_remove()

preview_frame.grid_rowconfigure(0, weight=1)
preview_frame.grid_columnconfigure(0, weight=1)

preview_tree = ttk.Treeview(preview_frame, height=12)
preview_tree["show"] = "tree"
preview_tree.tag_configure("new", foreground="blue")

preview_tree.column("#0", width=300, stretch=True)
preview_tree["show"] = "tree"

preview_scroll_y = tk.Scrollbar(preview_frame, orient="vertical", command=preview_tree.yview)
preview_scroll_x = tk.Scrollbar(preview_frame, orient="horizontal", command=preview_tree.xview)


preview_tree.configure(
    yscrollcommand=preview_scroll_y.set,
    xscrollcommand=preview_scroll_x.set
)

preview_tree.grid(row=0, column=0, sticky="nsew")
preview_scroll_y.grid(row=0, column=1, sticky="ns")
preview_scroll_x.grid(row=1, column=0, sticky="ew")

preview_mode_btn = tk.Button(
    mode_frame,
    text="Preview",
    width=8,
    relief="sunken",
    command=lambda: root.after(100, lambda: switch_create_mode("preview"))
)

preview_mode_btn.pack(side="left")

info_mode_btn = tk.Button(
    mode_frame,
    text="Info",
    width=8,
    command=lambda: switch_create_mode("info")
)

info_mode_btn.pack(side="left")

preview_frame.grid(row=0, column=0, sticky="nsew")
content_box.pack_forget()

content_box.insert("1.0",
"""print("Germany")
x = int(input("Enter number"))

if x > 0:
    print(f"{x} more than 0")
elif x < 0:
    print(f"{x} less than 0")
else:
    print(f"{x} equal 0")
""")






#delete GUI
delete_frame = tk.Frame(stack_frame)

left_delete = tk.Frame(delete_frame)
left_delete.pack(side="left", padx=40, pady=20, anchor="n")

tk.Label(left_delete, text="Destination:").pack(anchor="w")

delete_row = tk.Frame(left_delete)
delete_row.pack(anchor="w", pady=5)

delete_dest_entry = tk.Entry(delete_row, width=20)
delete_dest_entry.pack(side="left")

def choose_delete_folder():
    folder = filedialog.askdirectory()
    if folder:
        delete_dest_entry.delete(0, tk.END)
        delete_dest_entry.insert(0, folder)
        set_status("Delete folder selected")
    search_files()

delete_folder_btn = tk.Button(
    delete_row,
    text="<<",
    width=4,
    bg="#ffffff",
    relief="raised",
    cursor="hand2",
    command=choose_delete_folder,
)

delete_folder_btn.pack(side="left", padx=(5,0))

tk.Label(left_delete, text="Keywords:").pack(anchor="w")

delete_keyword_entry = tk.Entry(left_delete, width=25)
delete_keyword_entry.insert(0, "*")
delete_keyword_entry.pack(anchor="w", pady=5)

tk.Label(left_delete, text="File format:").pack(anchor="w")

delete_ext_entry = tk.Entry(left_delete, width=10)
delete_ext_entry.insert(0, ".py")
delete_ext_entry.pack(anchor="w", pady=5)

delete_keyword_entry.bind("<KeyRelease>", search_files)
delete_ext_entry.bind("<KeyRelease>", search_files)
delete_dest_entry.bind("<KeyRelease>", search_files)


delete_exec_btn = ttk.Button(
    left_delete,
    text="Delete",
    command=delete_found_files
)

delete_exec_btn.pack(pady=10)

right_delete = tk.Frame(delete_frame)
right_delete.pack(side="left", padx=40, pady=20, fill="both", expand=True)

right_delete.config(width=350, height=300)

delete_container = tk.Frame(right_delete)
delete_container.pack(fill="both", expand=True)

delete_container.grid_rowconfigure(0, weight=1)
delete_container.grid_columnconfigure(0, weight=1)

delete_tree = ttk.Treeview(
    delete_container,
    show="tree",
    selectmode="extended"
)
delete_tree.bind("<Button-1>", lambda e: delete_tree.focus_set())
delete_tree.bind("<KeyPress-Delete>", start_delete_hold)
delete_tree.bind("<KeyRelease-Delete>", stop_delete_hold)

delete_tree["show"] = "tree"
delete_tree.tag_configure("delete", foreground="red")
delete_scroll_y = tk.Scrollbar(delete_container, orient="vertical", command=delete_tree.yview)
delete_scroll_x = tk.Scrollbar(delete_container, orient="horizontal", command=delete_tree.xview)

delete_tree.configure(
    yscrollcommand=delete_scroll_y.set,
    xscrollcommand=delete_scroll_x.set
)
delete_tree.grid(row=0, column=0, sticky="nsew")
delete_scroll_y.grid(row=0, column=1, sticky="ns")
delete_scroll_x.grid(row=1, column=0, sticky="ew")
delete_container.grid_rowconfigure(1, minsize=0)
delete_tree.column("#0", width=300, stretch=False)

def delete_key_handler(event):
    items = delete_tree.selection()

    if not items:
        items = delete_tree.get_children()

    delete_found_files_custom(items)
    

root.bind("<Delete>", delete_key_handler)



# scrollbars
dest_scrollbar = tk.Scrollbar(dest_frame, orient="vertical", command=dest_tree.yview)
dest_scrollbar_x = tk.Scrollbar(dest_frame, orient="horizontal", command=dest_tree.xview)

# connect scrollbars
dest_tree.configure(
    yscrollcommand=dest_scrollbar.set,
    xscrollcommand=dest_scrollbar_x.set
)

# layout
dest_tree.grid(row=0, column=0, sticky="nsew")
dest_scrollbar.grid(row=0, column=1, sticky="ns")
dest_scrollbar_x.grid(row=1, column=0, sticky="ew")

dest_frame.grid_rowconfigure(0, weight=1)
dest_frame.grid_columnconfigure(0, weight=1)

# tree settings
dest_tree.column("#0", width=220, minwidth=220, stretch=False)
dest_tree.heading("#0", text="")
dest_tree.tag_configure("highlight", background="#cce8ff")

# context menu
dest_menu = tk.Menu(root, tearoff=0)

dest_menu.add_command(
    label="Open folder",
    command=lambda: os.startfile(destination_folder)
)

dest_menu.add_command(
    label="Clear",
    command=clear_destination
)

dest_menu.add_command(
    label="Copy path",
    command=lambda: root.clipboard_append(destination_folder)
)

# bindings
files_tree.bind("<MouseWheel>", lambda e: files_tree.yview_scroll(int(-1*(e.delta/120)), "units"))
dest_tree.bind("<MouseWheel>", lambda e: dest_tree.yview_scroll(int(-1*(e.delta/120)), "units"))

dest_tree.bind("<Motion>", dest_cursor_hover)
dest_tree.bind("<Leave>", lambda e: dest_tree.config(cursor=""))
dest_tree.bind("<Delete>", delete_destination_key)

# buttons
dest_buttons_frame = tk.Frame(center_frame)
dest_buttons_frame.grid(row=2, column=1, pady=(10,15))
dest_buttons_frame.grid_columnconfigure(0, weight=1)
dest_buttons_frame.grid_columnconfigure(1, weight=1)


choose_dest_btn = ttk.Button(
    dest_buttons_frame,
    text="Choose Folder",
    command=choose_destination
)
choose_dest_btn.grid(row=0, column=0, padx=5, sticky="e")

delete_dest_btn = ttk.Button(
    dest_buttons_frame,
    text="Clear",
    command=clear_destination
)
delete_dest_btn.grid(row=0, column=1, padx=5, sticky="w")

action_btn = ttk.Button(
    center_frame,
    text="MOVE",
    width=14,
    command=execute_action
)

action_btn.grid(row=3, column=0, columnspan=2, pady=(0,15))

#status bar (in work)
bottom_frame = tk.Frame(root, bd=1, relief="solid")
bottom_frame.pack(side="bottom", fill="x", pady=(8,0))

progress = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
progress.pack(fill=tk.X, padx=5, pady=(6,6))

status_label = tk.Label(
    bottom_frame,
    text="Ready",
    fg="black",
    bd=1,
    relief=tk.SUNKEN,
    anchor="w"
)
status_label.pack(fill=tk.X)

set_mode("MOVE")
switch_create_mode("preview")
root.mainloop()
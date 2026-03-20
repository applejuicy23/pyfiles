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
import re
from datetime import datetime
import webbrowser

#var
remove_hold_job = None
icons_cache = []
selected_files = []
selected_files_set = set()
destination_folder = ""
drag_select_start = None
current_mode = "MOVE"
active_button = None
delete_files_list = []
current_theme = "default"

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
        root.after(5000, reset_progress)

    except Exception as e:
        show_error(str(e))
#animation
def animate_progress(target):
    global progress_job

    current = progress["value"]

    if current >= target:
        return

    progress["value"] += 0.3

    progress_job = root.after(10, animate_progress, target)
def reset_progress():
    global progress_job

    if progress_job:
        root.after_cancel(progress_job)
        progress_job = None

    progress["value"] = 0
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
        root.after(5000, reset_progress)

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
        btn.config(
    relief="flat",
    bg=THEMES[current_theme]["panel"],
    fg=THEMES[current_theme]["fg"]
)

    if mode == "MOVE":
        move_btn.config(relief="flat", bg=THEMES[current_theme]["accent"], fg=THEMES[current_theme]["fg"], text="✔ MOVE")
        active_button = move_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "DUPLICATE":
        duplicate_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text="✔ COPY")
        active_button = duplicate_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "CREATE":
        create_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text="✔ CREATE")
        active_button = create_btn
        create_frame.pack(fill="both", expand=True)

    elif mode == "DELETE":
        delete_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text="✔ DELETE")
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

    root.after(5000, lambda: set_status("Ready"))
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
    remove_hold_job = root.after(500, lambda: delete_files(all=True))
def stop_delete_hold(event):
    global remove_hold_job
    if remove_hold_job:
        root.after_cancel(remove_hold_job)
        remove_hold_job = None

def start_drag_select(event):
    global drag_select_start
    tree = event.widget
    drag_select_start = tree.identify_row(event.y)
def drag_select(event):
    global drag_select_start

    tree = event.widget

    if not drag_select_start:
        return

    current = tree.identify_row(event.y)

    if not current:
        return

    start_index = tree.index(drag_select_start)
    current_index = tree.index(current)

    tree.selection_remove(tree.selection())

    low = min(start_index, current_index)
    high = max(start_index, current_index)

    items = tree.get_children()

    for i in range(low, high + 1):
        tree.selection_add(items[i])
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

def get_selected_delete_path():
    selection = delete_tree.selection()
    if not selection:
        return None

    index = delete_tree.index(selection[0])

    if index < len(delete_files_list):
        return delete_files_list[index]

    return None


def open_delete_file():
    path = get_selected_delete_path()
    if path:
        os.startfile(path)


def open_delete_folder():
    path = get_selected_delete_path()
    if path:
        os.startfile(os.path.dirname(path))


def copy_delete_path():
    path = get_selected_delete_path()
    if path:
        root.clipboard_clear()
        root.clipboard_append(path)

def show_delete_menu(event):
    item = delete_tree.identify_row(event.y)

    if item:
        delete_tree.selection_set(item)
        delete_menu.post(event.x_root, event.y_root)

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
def select_all(event):
    event.widget.selection_set(event.widget.get_children())
    return "break"

def switch_create_mode(mode):

    preview_mode_btn.config(relief="raised")
    info_mode_btn.config(relief="raised")

    if mode == "preview":
        preview_mode_btn.config(relief="sunken")

        update_preview()


        preview_frame.grid(row=0, column=0, sticky="nsew")


        content_box.grid_remove()
        content_scroll.grid_remove()
        content_scroll_x.grid_remove()


        preview_scroll_x.grid()
        preview_scroll_y.grid()

    elif mode == "info":
        info_mode_btn.config(relief="sunken")


        preview_frame.grid_remove()
        preview_scroll_x.grid_remove()
        preview_scroll_y.grid_remove()


        content_box.grid(row=0, column=0, sticky="nsew")
        content_scroll.grid(row=0, column=1, sticky="ns")
        content_scroll_x.grid(row=1, column=0, sticky="ew")


import webbrowser

def show_about():
    win = tk.Toplevel(root)
    win.title("About PyFiles")
    win.geometry("300x200")
    win.config(bg="#1e1e1e")  # под твой dark

    tk.Label(
        win,
        text="PyFiles v0.9.2rc\n\nNot-a-Simple GUI tool\nPython 3.14 + tkinter\nv0.9.2rc - RC1",
        bg="#1e1e1e",
        fg="white",
        justify="center"
    ).pack(pady=10)

    link = tk.Label(
        win,
        text="steam: id/applejuicy23",
        fg="cyan",
        bg="#1e1e1e",
        cursor="hand2"
    )
    link.pack()

    link.bind(
        "<Button-1>",
        lambda e: webbrowser.open("https://steamcommunity.com/id/applejuicy23")
    )


def show_whats_new():
    win = tk.Toplevel(root)
    win.title("What's New?")
    win.geometry("350x250")
    win.config(bg="#1e1e1e")  # под dark

    tk.Label(
        win,
        text="PyFiles v0.9.2rc",
        bg="#1e1e1e",
        fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=(10, 5))

    text = tk.Text(
        win,
        bg="#1e1e1e",
        fg="white",
        bd=0,
        highlightthickness=0
    )
    text.pack(expand=True, fill="both", padx=10, pady=10)

    text.insert("1.0",
        ">> new UI updated\n"
        ">> bug with width window is fixed\n"
        ">> dark mode is fixed (now it's all black)\n"
        ">> now you can change screen size\n"
        ">> in Delete mode added refresh button\n"
        ">> Ctrl+A & drag added for Delete mode\n"
        ">> Button About and What's new UI changed (preview)\n"
    )

    text.config(state="disabled")

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

def delete_files(all=False):
    if not delete_files_list:
        show_error("No files found")
        return

    if all:
        targets = delete_files_list.copy()
    else:
        selected = delete_tree.selection()

        if not selected:
            show_error("No files selected")
            return

        targets = []
        for item in selected:
            index = delete_tree.index(item)
            if index < len(delete_files_list):
                targets.append(delete_files_list[index])

    progress["maximum"] = len(targets)
    progress["value"] = 0

    for i, path in enumerate(targets, start=1):
        try:
            os.remove(path)
        except:
            pass

        animate_progress(i)
        root.update()

    search_files()
    set_success("Files deleted")
    root.after(5000, reset_progress)

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

    progress["maximum"] = count
    progress["value"] = 0

    for i, _ in enumerate(range(count), start=1):
        new_name = apply_prefix(prefix, name, i, ext)
        filename = f"{new_name}{ext}"
        path = os.path.join(folder, filename)

        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

        animate_progress(i)
        root.update()

    set_success("Files created")
    clear_all_files()
    update_preview()
    root.after(5000, reset_progress)

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

    delete_tree.delete(*delete_tree.get_children())
    delete_files_list.clear()

    if not folder or not os.path.isdir(folder):
        return

    try:
        files = sorted(os.listdir(folder), key=natural_sort_key)
    except:
        return

    for file in files:
        if keyword and keyword != "*":
            if keyword.lower() not in file.lower():
                continue

        if ext:
            if not file.endswith(ext):
                continue

        full_path = os.path.join(folder, file)

        delete_files_list.append(full_path)
        delete_tree.insert("", "end", text=file, tags=("delete",))

    adjust_delete_column_full()
    update_delete_scroll()

def add_placeholder(entry, placeholder):
    entry.insert(0, placeholder)
    entry.placeholder = placeholder
    entry.is_placeholder = True
    entry.config(fg="gray")

    def on_focus_in(event):
        if entry.is_placeholder:
            entry.delete(0, tk.END)
            entry.config(fg=THEMES[current_theme]["fg"])
            entry.is_placeholder = False

    def on_focus_out(event):
        if not entry.get():
            entry.insert(0, placeholder)
            entry.config(fg="gray")
            entry.is_placeholder = True

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

    for i in range(1, count + 1):
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
def natural_sort_key(s):
    return [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r'(\d+)', s)
    ]


#if error appeared
def show_error(msg):
    status_label.config(text=f"Error: {msg}", fg="red")

def set_status(msg):
    status_label.config(text=msg, fg=THEMES[current_theme]["fg"])

def set_success(msg):
    status_label.config(text=msg, fg="darkgreen")


#better to sort it but too much shit, later...
#window
root = TkinterDnD.Tk()
root.title("PyFiles v0.9.2rc")
root.geometry("700x600")

root.iconbitmap(resource_path("pyfiles2.ico"))
style = ttk.Style()
THEMES = {
    "default": {
        "bg": "SystemButtonFace",
        "fg": "black",
        "panel": "SystemButtonFace",
        "accent": "#e8e8e8",
        "scroll": "SystemButtonFace"
    },
    "light": {
        "bg": "#ffffff",
        "fg": "#000000",
        "panel": "#ffffff",
        "accent": "#eaeaea",
    },
    "dark": {
        "bg": "#1e1e1e",
        "fg": "#ffffff",
        "panel": "#2a2a2a",
        "accent": "#3a3a3a",
        "scroll": "#1e1e1e"
    }
}

def apply_theme(name):
    global current_theme
    current_theme = name
    t = THEMES[name]


    root.config(bg=t["bg"])
    main_frame.config(bg=t["bg"])
    content_frame.config(bg=t["bg"])
    stack_frame.config(bg=t["bg"])
    menu_frame.config(bg=t["panel"])
    title_frame.config(bg=t["panel"])
    title.config(bg=t["panel"], fg=t["fg"])
    icon_label.config(bg=t["panel"])
    entry_bg = "#ffffff" if name != "dark" else "#2a2a2a"
    entry_fg = "#000000" if name != "dark" else "#ffffff"

    for frame in (
    dest_row,
    delete_row,
    files_buttons_frame,
    dest_buttons_frame,
    mode_frame,
    content_container,
    delete_container,
    files_frame,
    dest_frame,
    preview_frame,
    delete_container,
    ):
        try:
            frame.config(bg=t["bg"])
        except:
            pass

    for sb in (
    files_scrollbar, files_scrollbar_x,
    dest_scrollbar, dest_scrollbar_x,
    content_scroll, content_scroll_x,
    preview_scroll_x, preview_scroll_y,
    delete_scroll_x, delete_scroll_y
):
        try:
            sb.config(
                bg=t["bg"],
                troughcolor=t["bg"],
                activebackground=t["accent"],
                highlightbackground=t["bg"],
                highlightcolor=t["bg"],
                highlightthickness=0,
                bd=0,
                relief="flat"
            )
        except:
            pass
    root.option_add("*Scrollbar.background", t["bg"])
    root.option_add("*Scrollbar.troughColor", t["bg"])
    root.option_add("*Scrollbar.activeBackground", t["accent"])


    if current_theme == "dark":
        style.configure("Vertical.TScrollbar",
            background="#2a2a2a",
            troughcolor="#1e1e1e",
            arrowcolor="#ffffff"
        )

        style.configure("Horizontal.TScrollbar",
            background="#2a2a2a",
            troughcolor="#1e1e1e",
            arrowcolor="#ffffff"
        )

    elif current_theme == "light":
        style.configure("Vertical.TScrollbar",
            background="#f2f2f2",
            troughcolor="#e6e6e6",
            arrowcolor="#000000"
        )

        style.configure("Horizontal.TScrollbar",
            background="#f2f2f2",
            troughcolor="#e6e6e6",
            arrowcolor="#000000"
        )

    else:
        style.configure("Vertical.TScrollbar",
            background="#e8e8e8",
            troughcolor="#e8e8e8",
            arrowcolor="#000000"
        )

        style.configure("Horizontal.TScrollbar",
            background="#e8e8e8",
            troughcolor="#e8e8e8",
            arrowcolor="#000000"
        )

    style.map("Horizontal.TScrollbar",
        background=[("active", t["accent"])]
    )
    style.map("Vertical.TScrollbar",
        background=[("active", t["accent"])]
    )
    

    for entry in (
        file_name_entry,
        format_entry,
        prefix_entry,
        count_entry,
        dest_entry,
        delete_dest_entry,
        delete_keyword_entry,
        delete_ext_entry
    ):
        entry.config(
            bg=entry_bg,
            insertbackground=entry_fg
        )

        if getattr(entry, "is_placeholder", False):
            entry.config(fg="gray")
        else:
            entry.config(fg=entry_fg)

    for lbl in (
    files_label, dest_label
):
        lbl.config(bg=t["bg"], fg=t["fg"])
    
    for widget in right_create.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])
    for lbl in (theme_light, theme_gray, theme_dark):
        lbl.bind("<Enter>", theme_hover)
        lbl.bind("<Leave>", theme_leave)

    content_box.config(
        bg=entry_bg,
        fg=entry_fg,
        insertbackground=entry_fg
    )


    if current_theme == "dark":
        style.configure("TButton",
            background="#2a2a2a",
            foreground="#ffffff",
            padding=6,
            borderwidth=1,
            relief="sunken"
        )
    else:
        style.configure("TButton",
            background=t["panel"],
            foreground=t["fg"],
            padding=6,
            borderwidth=1,
            relief="sunken"
        )

    if current_theme == "dark":
        preview_mode_btn.config(bg="#2a2a2a", fg="#ffffff")
        info_mode_btn.config(bg="#2a2a2a", fg="#ffffff")
    else:
        preview_mode_btn.config(bg=t["panel"], fg=t["fg"])
        info_mode_btn.config(bg=t["panel"], fg=t["fg"])

    for btn in (move_btn, duplicate_btn, create_btn, delete_btn,
            about_tab, whatsnew_tab,
            theme_light, theme_gray, theme_dark):
        btn.config(fg=t["fg"])
    for lbl in (theme_light, theme_gray, theme_dark):
        lbl.config(bg=t["bg"], fg=t["fg"])



    left_panel.config(bg=t["panel"])
    bottom_frame.config(bg=t["panel"])


    for btn in (move_btn, duplicate_btn, create_btn, delete_btn):
        btn.config(bg=t["panel"], fg=t["fg"], activebackground=t["accent"])


    about_tab.config(bg=t["panel"], fg=t["fg"])
    whatsnew_tab.config(bg=t["panel"], fg=t["fg"])
    if current_theme == "dark":
        about_tab.config(bg="#e0e0e0", fg="#000000")
        whatsnew_tab.config(bg="#e0e0e0", fg="#000000")
    else:
        about_tab.config(bg=t["panel"], fg=t["fg"])
        whatsnew_tab.config(bg=t["panel"], fg=t["fg"])

    status_label.config(bg=t["panel"], fg=t["fg"])


    for f in (
        center_frame, create_frame, delete_frame,
        left_create, right_create,
        left_delete, right_delete,
        dest_row, delete_row,
        files_buttons_frame, dest_buttons_frame,
        mode_frame, content_container, delete_container
    ):
        try:
            f.config(bg=t["bg"])
        except:
            pass
    
        content_box.config(
        bg=entry_bg,
        fg=entry_fg,
        insertbackground=entry_fg,
        highlightbackground=t["bg"],
        highlightcolor=t["bg"]
    )
    #ttk
    style.configure("Treeview",
        background=entry_bg,
        fieldbackground=entry_bg,
        foreground=entry_fg
    )

    style.configure("TButton",
        background=t["panel"],
        foreground=t["fg"]
    )

    style.configure("TProgressbar",
        troughcolor=t["panel"],
        background="#cac4c4"
    )

    style.configure("TButton",
    background=t["panel"],
    foreground=t["fg"]
    )
    style.map("TButton",
        background=[("active", t["accent"])],
        foreground=[("!disabled", "#ffffff" if current_theme == "dark" else t["fg"])]
    )

    #scrollbar

    folder_btn.config(
        highlightbackground=t["bg"],
        activebackground=t["accent"]
    )

    delete_folder_btn.config(
        highlightbackground=t["bg"],
        activebackground=t["accent"]
    )

    for btn in (
        choose_files_btn, delete_file_btn,
        choose_dest_btn, delete_dest_btn,
        action_btn, create_btn_exec,
        delete_exec_btn
    ):
        try:
            btn.configure(style="TButton")
        except:
            pass

    for widget in right_create.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])

    for widget in left_create.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])

    for widget in left_delete.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])

    if current_theme == "dark":
        about_tab.config(bg="#2a2a2a", fg="#ffffff")
        whatsnew_tab.config(bg="#2a2a2a", fg="#ffffff")
    else:
        about_tab.config(bg=t["panel"], fg=t["fg"])
        whatsnew_tab.config(bg=t["panel"], fg=t["fg"])

    if current_theme == "dark":
        btn_bg = "#2a2a2a"
        btn_fg = "#ffffff"
    else:
        btn_bg = "#e0e0e0" 
        btn_fg = "#000000" 


    for btn in (folder_btn, delete_folder_btn):
        btn.config(
            bg=btn_bg,
            fg=btn_fg,
            activebackground="#cccccc" if current_theme != "dark" else "#3a3a3a",
            activeforeground=btn_fg,
            relief="raised",
            bd=1
        )
    style.configure("Horizontal.TProgressbar",
    background="#4caf50"
    )
    if current_theme == "light":
        for btn in (folder_btn, delete_folder_btn):
            btn.config(
                bg="#ffffff", 
                fg="#000000", 
                activebackground="#eaeaea",
                activeforeground="#000000",
                relief="raised",
                bd=1
            )

    update_preview_colors()
    apply_preview_theme()
    


def update_preview_colors():
    if current_theme == "dark":
        preview_tree.tag_configure("new", foreground="#ffffff")
    else:
        preview_tree.tag_configure("new", foreground="blue")

def apply_preview_theme():
    t = THEMES[current_theme]

    entry_bg = "#ffffff" if current_theme != "dark" else "#2a2a2a"
    entry_fg = "#000000" if current_theme != "dark" else "#ffffff"

    style.configure("Preview.Treeview",
        background=entry_bg,
        fieldbackground=entry_bg,
        foreground=entry_fg
    )

    preview_frame.config(bg=t["bg"])
    content_container.config(bg=t["bg"])

def theme_hover(e):
    e.widget.config(bg=THEMES[current_theme]["accent"])

def theme_leave(e):
    e.widget.config(bg=THEMES[current_theme]["bg"])


style.theme_use("clam")
icon_img = ImageTk.PhotoImage(
    Image.open(resource_path("pyfiles2.ico")).resize((16, 16))
)
root.icon_img = icon_img

file_menu = tk.Menu(root, tearoff=0)

file_menu.add_command(label="Open", command=lambda: open_selected_file(None))
file_menu.add_command(label="Remove", command=delete_selected_file)
file_menu.add_separator()
file_menu.add_command(label="Reveal in Explorer", command=lambda: os.startfile(os.path.dirname(selected_files[files_tree.index(files_tree.selection()[0])])))
file_menu.add_command(label="Copy path", command=lambda: root.clipboard_append(selected_files[files_tree.index(files_tree.selection()[0])]))

delete_menu = tk.Menu(root, tearoff=0)

delete_menu.add_command(label="Open", command=lambda: open_delete_file())
delete_menu.add_command(label="Remove", command=lambda: delete_files())
delete_menu.add_separator()
delete_menu.add_command(label="Reveal in Explorer", command=lambda: open_delete_folder())
delete_menu.add_command(label="Copy path", command=lambda: copy_delete_path())
        
#buttons setup
menu_frame = tk.Frame(root)
menu_frame.pack(fill="x")

title_frame = tk.Frame(menu_frame)
title_frame.pack(side="left", padx=10)

icon_label = tk.Label(title_frame, image=icon_img)
icon_label.pack(side="left")

title = tk.Label(
    title_frame,
    text=" PyFiles v0.9.2rc",
    font=("Segoe UI", 10)
)
title.pack(side="left")

tabs_frame = tk.Frame(menu_frame)
tabs_frame.pack(side="left", padx=10)

def select_tab(tab):
    tab.config(bg="white")

    def reset():
        if current_theme == "dark":
            tab.config(bg="#2a2a2a", fg="#ffffff")
        else:
            t = THEMES[current_theme]
            tab.config(bg=t["panel"], fg=t["fg"])

    root.after(150, reset)






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

about_tab = tk.Label(
    left_panel,
    text="ABOUT",
    padx=12,
    pady=3,
    bd=1,
    relief="raised",
    cursor="hand2"
)

about_tab.pack(padx=0, pady=(20,0))

whatsnew_tab = tk.Label(
    left_panel,
    text="WHAT'S NEW",
    padx=12,
    pady=3,
    bd=1,
    relief="raised",
    cursor="hand2"
)

whatsnew_tab.pack(padx=0, pady=(5,0))

theme_light = tk.Label(left_panel, text="Light", cursor="hand2")
theme_light.pack(pady=(10,0))

theme_gray = tk.Label(left_panel, text="Default", cursor="hand2")
theme_gray.pack()

theme_dark = tk.Label(left_panel, text="Dark", cursor="hand2")
theme_dark.pack()

theme_light.bind("<Button-1>", lambda e: apply_theme("light"))
theme_gray.bind("<Button-1>", lambda e: apply_theme("default"))
theme_dark.bind("<Button-1>", lambda e: apply_theme("dark"))
about_tab.bind("<Button-1>", lambda e: (select_tab(about_tab), show_about()))
whatsnew_tab.bind("<Button-1>", lambda e: (select_tab(whatsnew_tab), show_whats_new()))

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

def sidebar_hover(event):
    if event.widget != active_button:
        event.widget.config(bg=THEMES[current_theme]["accent"])

def sidebar_leave(event):
    if event.widget != active_button:
        event.widget.config(bg=THEMES[current_theme]["panel"])

for btn in (move_btn, duplicate_btn, create_btn, delete_btn):
    btn.bind("<Enter>", sidebar_hover)
    btn.bind("<Leave>", sidebar_leave)


if active_button:
    active_button.config(
        bg=THEMES[current_theme]["accent"],
        fg=THEMES[current_theme]["fg"]
    )


#files column


files_label = tk.Label(center_frame, text="Source Files")
files_label.grid(row=0, column=0, pady=(10,5))

files_frame = tk.Frame(center_frame)
files_frame.grid(row=1, column=0, padx=20, sticky="nsew")
files_frame.config(width=260)
files_frame.grid_propagate(False)

files_tree = ttk.Treeview(
    files_frame,
    show="tree",
    height=10,
    selectmode="extended",
)

files_tree["show"] = "tree"

files_scrollbar = ttk.Scrollbar(files_frame, orient="vertical", command=files_tree.yview)
files_scrollbar_x = ttk.Scrollbar(files_frame, orient="horizontal", command=files_tree.xview)

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
dest_frame.config(width=260)
dest_frame.grid_propagate(False)
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
    relief="sunken",
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
right_create.pack(side="left", padx=40, pady=20, fill="both", expand=True)
right_create.pack_propagate(False)
tk.Label(
    right_create,
    text="Contain information inside it:"
).pack(anchor="w")

mode_frame = tk.Frame(right_create)
mode_frame.pack(anchor="w", pady=(0,5))

content_container = tk.Frame(right_create)
content_container.pack_propagate(False)
content_container.pack(fill="both", expand=True)
content_container.grid_rowconfigure(0, weight=1)
content_container.grid_rowconfigure(1, weight=0)
content_container.grid_columnconfigure(0, weight=1)

preview_frame = tk.Frame(content_container)


right_create.pack(side="left", padx=40, pady=20, fill="both", expand=True)
right_create.config(width=350, height=300)

content_box = tk.Text(content_container)
content_box.config(
    font=("Consolas", 10),
    insertbackground="black",
    undo=True
)
content_box.config(tabs=("1c"))

content_scroll = ttk.Scrollbar(content_container, orient="vertical", command=content_box.yview)
content_box.configure(yscrollcommand=content_scroll.set)

preview_frame.grid(row=0, column=0, sticky="nsew")
content_box.grid_remove()
content_scroll.grid_remove()

preview_frame.grid_rowconfigure(0, weight=1)
preview_frame.grid_columnconfigure(0, weight=1)

preview_tree = ttk.Treeview(preview_frame, height=12, style="Preview.Treeview")
preview_tree["show"] = "tree"
preview_tree.tag_configure("new", foreground="blue")

apply_preview_theme()

preview_tree.column("#0", width=300, stretch=True)
preview_tree["show"] = "tree"

preview_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=preview_tree.yview)
preview_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=preview_tree.xview)


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

content_scroll_x = ttk.Scrollbar(content_container, orient="horizontal", command=content_box.xview)

content_box.configure(
    xscrollcommand=content_scroll_x.set,
    wrap="none"
)

content_scroll_x.grid(row=1, column=0, sticky="ew")

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

add_placeholder(delete_dest_entry, "Destination folder path")

def choose_delete_folder():
    folder = filedialog.askdirectory()
    if folder:
        delete_dest_entry.delete(0, tk.END)
        delete_dest_entry.insert(0, folder)
        delete_dest_entry.is_placeholder = False
        delete_dest_entry.config(fg=THEMES[current_theme]["fg"])
        set_status("Delete folder selected")
    search_files()

delete_folder_btn = tk.Button(
    delete_row,
    text="<<",
    width=4,
    bg="#ffffff",
    relief="sunken",
    cursor="hand2",
    command=choose_delete_folder,
)

delete_folder_btn.pack(side="left", padx=(5,0))

tk.Label(left_delete, text="Keywords:").pack(anchor="w")

delete_keyword_entry = tk.Entry(left_delete, width=25)
delete_keyword_entry.pack(anchor="w", pady=5)
add_placeholder(delete_keyword_entry, "*")

tk.Label(left_delete, text="File format:").pack(anchor="w")

delete_ext_entry = tk.Entry(left_delete, width=10)
delete_ext_entry.pack(anchor="w", pady=5)
add_placeholder(delete_ext_entry, ".py")


delete_keyword_entry.bind("<KeyRelease>", search_files)
delete_ext_entry.bind("<KeyRelease>", search_files)
delete_dest_entry.bind("<KeyRelease>", search_files)


delete_exec_btn = ttk.Button(
    left_delete,
    text="Delete",
    command=lambda: delete_files()
)


delete_exec_btn.pack(pady=10)

delete_exec_btn.bind("<ButtonPress-1>", start_delete_hold)
delete_exec_btn.bind("<ButtonRelease-1>", stop_delete_hold)

right_delete = tk.Frame(delete_frame)
right_delete.pack(side="left", padx=40, pady=20, fill="both", expand=True)

right_delete.config(width=350, height=300)

refresh_btn = ttk.Button(
    right_delete,
    text="Refresh",
    command=search_files
)

refresh_btn.pack(anchor="ne", pady=(0,5))

delete_container = tk.Frame(right_delete)
delete_container.pack(fill="both", expand=True)

delete_container.grid_rowconfigure(0, weight=1)
delete_container.grid_columnconfigure(0, weight=1)

delete_tree = ttk.Treeview(
    delete_container,
    show="tree",
    selectmode="extended"
)
delete_tree.bind("<Delete>", lambda e: delete_files())
delete_tree.bind("<Button-3>", show_delete_menu)
delete_tree.bind("<Button-1>", start_drag_select)
delete_tree.bind("<B1-Motion>", drag_select)
delete_tree.bind("<ButtonRelease-1>", stop_drag_select)
delete_tree.bind("<Control-a>", select_all)
delete_tree.bind("<Control-A>", select_all)

delete_tree["show"] = "tree"
delete_tree.tag_configure("delete", foreground="red")
delete_scroll_y = ttk.Scrollbar(delete_container, orient="vertical", command=delete_tree.yview)
delete_scroll_x = ttk.Scrollbar(delete_container, orient="horizontal", command=delete_tree.xview)

delete_tree.configure(
    yscrollcommand=delete_scroll_y.set,
    xscrollcommand=delete_scroll_x.set
)
delete_tree.grid(row=0, column=0, sticky="nsew")
delete_scroll_y.grid(row=0, column=1, sticky="ns")
delete_scroll_x.grid(row=1, column=0, sticky="ew")
delete_container.grid_rowconfigure(1, minsize=0)
delete_tree.column("#0", width=100, stretch=False)
def delete_key_handler(event):
    delete_files()
    



# scrollbars
dest_scrollbar = ttk.Scrollbar(dest_frame, orient="vertical", command=dest_tree.yview)
dest_scrollbar_x = ttk.Scrollbar(dest_frame, orient="horizontal", command=dest_tree.xview)

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
apply_theme("default")
root.update_idletasks()
root.mainloop()
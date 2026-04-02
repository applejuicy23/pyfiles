import win32gui
import win32ui
import win32con
import random
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
from pathlib import Path
import winshell
import pythoncom
from send2trash import send2trash
from send2trash.win import legacy as send2trash_legacy
import win32com.client
import subprocess
import json
import tkinter.messagebox as mb
import uuid

#creating delete folder
BIN_DIR = os.path.join(os.getcwd(), ".pyfiles_bin")
META_FILE = os.path.join(BIN_DIR, "meta.json")
CACHE_DIR = os.path.join(os.getcwd(), "PyFiles_Cache")
os.makedirs(CACHE_DIR, exist_ok=True)
os.makedirs(BIN_DIR, exist_ok=True)

if not os.path.exists(META_FILE):
    with open(META_FILE, "w") as f:
        json.dump({}, f)

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
about_window = None
whatsnew_window = None
last_status = None
is_dragging = False
log_text = None
deleted_stack = []
deleted_files = []
icon_cache_map = {}
log_buffer = []
bin_items_map = {}
bin_icons_cache = []
ext_icon_cache = {}
preview_icon_cache = {}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
deleted_file_meta = {}
excluded_paths = set()
added_paths = set()


status_job = None

def show_status(msg, duration=5000, is_error=False):
    global status_job

    if status_job:
        root.after_cancel(status_job)

    color = "red" if is_error else THEMES[current_theme]["fg"]

    status_label.config(text=msg, fg=color)

    status_job = root.after(duration, lambda: set_status("Ready"))
def show_error_status(msg):
    show_status(f"Error: {msg}", duration=15000, is_error=True)
def format_items(files, folders, names=None, action="Added"):
    if names and len(names) == 1:
        name = names[0]
        if files:
            return f"{action} file {name}"
        else:
            return f"{action} folder {name}"

    parts = []
    if folders:
        parts.append(f"{folders} folder{'s' if folders != 1 else ''}")
    if files:
        parts.append(f"{files} file{'s' if files != 1 else ''}")

    return f"{action} {' and '.join(parts)}"
def remove_items(paths):
    files = 0
    folders = 0
    names = []
    for p in paths:
        added_paths.discard(norm(p))
    for p in paths:
        if os.path.isfile(p):
            files += 1
        elif os.path.isdir(p):
            folders += 1

        names.append(os.path.basename(p))

    msg = format_items(files, folders, names if len(paths)==1 else None, "Removed")
    show_status(msg)
def count_inside_folder(folder):
    files = 0
    folders = 0

    for root, dirs, filenames in os.walk(folder):
        folders += len(dirs)
        files += len(filenames)

    return folders, files


#up to release NEED TO MANAGE CODE, its hard to read and ducking my brain
#functions
#choosing files
def get_preview_icon(filename):
    ext = os.path.splitext(filename)[1].lower()

    # 💥 НЕТ РАСШИРЕНИЯ → ❓
    if not ext:
        return unknown_icon

    # 💥 ЕСТЬ РАСШИРЕНИЕ → стандартная логика (через твою функцию)
    return get_icon_by_extension(ext)
def choose_files():
    files = filedialog.askopenfilenames()
    files = root.tk.splitlist(files)

    if not files:
        return

    for f in files:
        name = os.path.basename(f)
        icon = get_file_icon(f)

        files_tree.insert(
            "",
            "end",
            text="\u2002" + name,
            image=icon if icon else "",
            values=(f,)
        )
        files_tree.update_idletasks()
        adjust_tree_column_full(files_tree)

        if icon:
            icons_cache.append(icon)

    names = [os.path.basename(f) for f in files]

    msg = format_items(len(files), 0, names if len(files)==1 else None)
    show_status(msg)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_shortcut_target(path):
    try:
        ext = os.path.splitext(path)[1].lower()


        if ext == ".lnk":
            shell_link = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell_link.CreateShortcut(path)
            return shortcut.Targetpath


        if ext == ".url":
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("IconFile="):
                        return line.strip().split("=", 1)[1]

        return None

    except Exception as e:
        print("shortcut error:", e)
        return None
def get_cached_preview_icon(ext):
    if not ext:
        return None

    if not ext.startswith("."):
        ext = "." + ext

    ext = ext.lower()

    if ext in preview_icon_cache:
        return preview_icon_cache[ext]

    icon = get_create_preview_icon(ext)

    if icon:
        preview_icon_cache[ext] = icon

    return icon
def get_folder_icon():
    try:
        flags = shellcon.SHGFI_ICON | shellcon.SHGFI_SMALLICON

        ret, info = shell.SHGetFileInfo(
            "C:\\Windows", 
            win32con.FILE_ATTRIBUTE_DIRECTORY,
            flags | shellcon.SHGFI_USEFILEATTRIBUTES
        )

        return info[0]

    except Exception as e:
        print("folder icon error:", e)
        return None

def get_file_icon(path):
    try:
        ext = os.path.splitext(path)[1].lower()

        if ext in IMAGE_EXTENSIONS:
            return image_icon


        real_path = get_shortcut_target(path)
        if real_path and os.path.exists(real_path):
            path = real_path
            ext = os.path.splitext(path)[1].lower()


        key = path

        if key in icon_cache_map:
            return icon_cache_map[key]

        flags = shellcon.SHGFI_ICON | shellcon.SHGFI_SMALLICON

        if os.path.isdir(path):
            hicon = get_folder_icon()

        else:

            if os.path.exists(path):
                ret, info = shell.SHGetFileInfo(
                    path,
                    win32con.FILE_ATTRIBUTE_NORMAL,
                    flags
                )
                hicon = info[0]
            else:
                hicon = None


            if not hicon:
                ret, info = shell.SHGetFileInfo(
                    path,
                    win32con.FILE_ATTRIBUTE_NORMAL,
                    flags | shellcon.SHGFI_USEFILEATTRIBUTES
                )
                hicon = info[0]



        if not hicon:
            ret, info = shell.SHGetFileInfo(
                path,
                win32con.FILE_ATTRIBUTE_NORMAL,
                shellcon.SHGFI_ICON |
                shellcon.SHGFI_SMALLICON |
                shellcon.SHGFI_USEFILEATTRIBUTES
            )
            hicon = info[0]



        if not hicon:
            ret, info = shell.SHGetFileInfo(
                path,
                0,
                shellcon.SHGFI_ICON |
                shellcon.SHGFI_SMALLICON |
                shellcon.SHGFI_SYSICONINDEX
            )
            hicon = info[0]


        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, 16, 16)

        hdc_mem = hdc.CreateCompatibleDC()
        hdc_mem.SelectObject(hbmp)

        win32gui.DrawIconEx(
            hdc_mem.GetHandleOutput(),
            0, 0,
            hicon, 16, 16,
            0, None,
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
        hdc_mem.DeleteDC()
        hdc.DeleteDC()

        icon = ImageTk.PhotoImage(img)

        if icon:
            icon_cache_map[key] = icon

        return icon

    except Exception as e:
        print("icon error:", e)
        return None
def get_icon_by_extension(ext):
    try:
        if not ext.startswith("."):
            ext = "." + ext

        fake_name = "file" + ext

        flags = (
            shellcon.SHGFI_ICON |
            shellcon.SHGFI_SMALLICON |
            shellcon.SHGFI_USEFILEATTRIBUTES
        )

        ret, info = shell.SHGetFileInfo(
            fake_name,
            win32con.FILE_ATTRIBUTE_NORMAL,
            flags
        )

        return info[0]

    except Exception as e:
        print("ext icon error:", e)
        return None
#choosing destination for files
def choose_destination():
    global destination_folder

    folder = filedialog.askdirectory()

    if not folder:
        return

    folder = os.path.normpath(folder)
    destination_folder = folder

    dest_tree.delete(*dest_tree.get_children())

    try:
        icon = get_file_icon(folder)
    except:
        icon = None

    dest_tree.insert(
        "",
        "end",
        text=folder,
        image=icon if icon else "",
        values=(folder,)
    )

    adjust_tree_column_full(dest_tree)

    if icon:
        icons_cache.append(icon)

    show_status("Destination selected")

def get_destination():
    items = dest_tree.get_children()
    if not items:
        return None

    return dest_tree.item(items[0], "values")[0]
#moving files
def count_inside_folder(folder_path):
    files = 0
    folders = 0

    for root_dir, dirs, filenames in os.walk(folder_path):
        folders += len(dirs)
        files += len(filenames)

    return folders, files
def move_to_cache(path, base_src):
    rel_path = os.path.relpath(path, base_src)
    target = os.path.join(CACHE_DIR, rel_path)

    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.move(path, target)
def move_to_destination(path, base_src, base_dst):
    rel_path = os.path.relpath(path, base_src)
    target = os.path.join(base_dst, rel_path)

    os.makedirs(os.path.dirname(target), exist_ok=True)

    if os.path.exists(target):
        os.remove(target)

    shutil.move(path, target)
def get_visible_paths(tree):
    visible = set()

    def collect(item):
        values = tree.item(item, "values")
        if values:
            visible.add(norm(values[0]))

        for child in tree.get_children(item):
            collect(child)

    for item in tree.get_children():
        collect(item)

    return visible
def is_inside_selected(path, selected_paths):
    p = norm(path)

    for sel in selected_paths:
        sel = norm(sel)
        if p.startswith(sel):
            return True

    return False
def is_visible(path, visible_paths):
    p = norm(path)

    for v in visible_paths:
        v = norm(v)

        if p == v or p.startswith(v + os.sep):
            return True

    return False
excluded_paths = set()
def mark_error_in_tree(tree, path):
    path = norm(path)

    def search(item):
        values = tree.item(item, "values")
        if values and norm(values[0]) == path:
            return item

        for child in tree.get_children(item):
            res = search(child)
            if res:
                return res

        return None

    for item in tree.get_children():
        res = search(item)
        if res:
            # 💥 красим сам файл
            tree.item(res, tags=("error",))

            # 💥 красим родителя (папку)
            parent = tree.parent(res)
            if parent:
                tree.item(parent, tags=("error",))

            return
def norm(p):
    return os.path.abspath(os.path.normcase(os.path.normpath(p)))
def remove_selected():
    selected = files_tree.selection()

    def collect_paths(item):
        paths = []

        values = files_tree.item(item, "values")
        if values:
            paths.append(values[0])

        for child in files_tree.get_children(item):
            paths.extend(collect_paths(child))

        return paths

    for item in selected:
        all_paths = collect_paths(item)

        for p in all_paths:
            excluded_paths.add(norm(p))  # 💥 теперь ВСЕ вложенные

        files_tree.delete(item)
def clean_tree(tree, error_paths):
    def process(item):
        values = tree.item(item, "values")

        keep = False

        if values:
            path = norm(values[0])
            if path in error_paths:
                keep = True

        # проверяем детей
        for child in tree.get_children(item):
            if process(child):
                keep = True

        if not keep:
            tree.delete(item)

        return keep

    for item in tree.get_children():
        process(item)
def is_excluded(path):
    p = norm(path)

    for ex in excluded_paths:
        if p == ex:
            return True

    return False
def get_visible_paths(tree):
    visible = set()

    def walk(item):
        values = tree.item(item, "values")
        if values:
            visible.add(norm(values[0]))

        for child in tree.get_children(item):
            walk(child)

    for item in tree.get_children():
        walk(item)

    return visible
def move_files():
    files_count = 0
    folders_count = 0
    destination = get_destination()
    paths = get_effective_selection(files_tree)

    # 💥 ВОТ СЮДА ПЕРЕНЕСТИ
    errors = []
    error_paths = set()
    logs = []

    all_tree_paths = get_all_tree_items(files_tree)

    for p in all_tree_paths:
        if not os.path.exists(p):
            err = f"{os.path.basename(p)} >> ERROR: File not found"

            errors.append(err)
            logs.append(err)
            log_to_console(err)

            mark_error_in_tree(files_tree, p)
            error_paths.add(norm(p))
    if not paths:
        for item in files_tree.get_children(""):
            values = files_tree.item(item, "values")
            if values:
                paths.append(values[0])

    if not paths or not destination:
        show_error("Select files and destination")
        return

    try:
        progress["maximum"] = len(paths)
        progress["value"] = 0
        start_console_log("MOVING FILES", len(paths))

        for i, path in enumerate(paths, start=1):
            try:
                if not os.path.exists(path):
                    raise FileNotFoundError("File not found")
                name = os.path.basename(path)
                target = os.path.join(destination, name)

                # 💥 конфликт имён
                if os.path.exists(target):
                    choice = mb.askyesnocancel(
                        "File exists",
                        f"{name} already exists.\n\nYes = Replace\nNo = Create copy\nCancel = Skip"
                    )

                    if choice is None:
                        continue
                    elif choice:
                        try:
                            if os.path.isdir(target):
                                shutil.rmtree(target)
                            else:
                                os.remove(target)
                        except:
                            pass
                    else:
                        base, ext = os.path.splitext(name)
                        counter = 1
                        while True:
                            new_name = f"{base} ({counter}){ext}"
                            new_target = os.path.join(destination, new_name)
                            if not os.path.exists(new_target):
                                target = new_target
                                break
                            counter += 1

                # 💥 ФАЙЛ
                if os.path.isfile(path):
                    files_count += 1

                    if is_excluded(path):
                        move_to_cache(path, os.path.dirname(path))
                    else:
                        shutil.move(path, target)

                # 💥 ПАПКА
                elif os.path.isdir(path):
                    root_target = os.path.join(destination, os.path.basename(path))
                    os.makedirs(root_target, exist_ok=True)

                    for root_dir, dirs, files in os.walk(path):
                        # файлы
                        for fname in files:
                            full = os.path.join(root_dir, fname)

                            # 💥 ДОБАВЛЕНА ПРОВЕРКА НА КЭШ
                            if is_excluded(full):
                                move_to_cache(full, os.path.dirname(full))
                                continue

                            if not os.path.exists(full):
                                err = f"{fname} >> ERROR: File not found"
                                errors.append(err)
                                logs.append(err)
                                log_to_console(err)

                                mark_error_in_tree(files_tree, full)
                                error_paths.add(norm(full))
                                continue

                            rel = os.path.relpath(full, path)
                            dest_file = os.path.join(root_target, rel)

                            os.makedirs(os.path.dirname(dest_file), exist_ok=True)

                            shutil.move(full, dest_file)
                            files_count += 1

                    # удаляем пустую папку
                    if not any(norm(p).startswith(norm(path)) for p in error_paths):
                        shutil.rmtree(path, ignore_errors=True)
                        folders_count += 1

                msg = f"File {name} moved ({destination})"
                logs.append(msg)
                log_to_console(msg)

            except Exception as e:
                err = f"File {os.path.basename(path)} >> ERROR: {str(e)}"
                errors.append(err)
                logs.append(err)
                log_to_console(err)
                mark_error_in_tree(files_tree, path)  # 💥 ВОТ ЭТО
                error_paths.add(norm(path))
            animate_progress(i)
            root.update()

        save_log(logs, errors)

        done = len(logs) - len(errors)
        total = len(logs)

        msg = f"Moved: {folders_count} folders, {files_count} files | Done: {done} | Errors: {len(errors)} | Total: {total} items"

        show_status(msg)

        root.after(5000, reset_progress)
        clean_tree(files_tree, error_paths)
        for p in error_paths:
            mark_error_in_tree(files_tree, p)
        excluded_paths.clear()
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
def clear_delete_selection(event=None):
    for item in delete_tree.selection():
        delete_tree.selection_remove(item)
def get_effective_selection(tree):
    selected = set(tree.selection())
    result = []

    def is_child_of_selected(item):
        parent = tree.parent(item)
        while parent:
            if parent in selected:
                return True
            parent = tree.parent(parent)
        return False

    for item in selected:
        if is_child_of_selected(item):
            continue

        values = tree.item(item, "values")
        if not values:
            continue

        result.append(values[0])

    return result
def open_selected_file(event):
    item = files_tree.selection()

    if not item:
        return

    item = item[0]

    path = files_tree.item(item, "values")[0]

    try:
        os.startfile(path)
    except Exception as e:
        show_error(str(e))

#duplicating files
def duplicate_files():
    if not destination_folder:
        show_error("Select destination")
        return
    try:
        selected = files_tree.selection()

        if not selected:
            files_to_copy = get_all_tree_items(files_tree)
        else:
            files_to_copy = [
                files_tree.item(item, "values")[0]
                for item in selected
                if files_tree.item(item, "values")
            ]

        if not files_to_copy:
            show_error("Nothing to copy")
            return

        progress["maximum"] = len(files_to_copy)
        progress["value"] = 0

        errors = []
        logs = []

        start_console_log("COPYING FILES", len(files_to_copy))

        for i, f in enumerate(files_to_copy, start=1):
            try:
                if not os.path.exists(f):
                    raise FileNotFoundError("File not found")
                if os.path.isfile(f):
                    shutil.copy2(f, destination_folder)

                elif os.path.isdir(f):
                    shutil.copytree(
                        f,
                        os.path.join(destination_folder, os.path.basename(f)),
                        dirs_exist_ok=True
                    )

                msg = f"File {os.path.basename(f)} copied ({destination_folder})"
                logs.append(msg)
                log_to_console(msg)

            except Exception as e:
                    # Безопасно получаем имя файла, если оно есть
                    filename = os.path.basename(f) if 'f' in locals() else "Unknown Target"
                    err = f"File {filename} >> ERROR: {str(e)}"
                    errors.append(err)
                    logs.append(err)
                    log_to_console(err)
                    
                    if 'f' in locals():
                        mark_error_in_tree(files_tree, f)

            animate_progress(i)
            root.update()

        save_log(logs, errors)
        
        done = len(logs) - len(errors)
        total = len(logs)

        msg = f"Copied: {total} items | Done: {done} | Errors: {len(errors)}"

        show_status(msg)
        root.after(5000, reset_progress)

    except Exception as e:
        err = f"File {os.path.basename(f)} >> ERROR: {str(e)}"
        errors.append(err)
        logs.append(err)
        log_to_console(err)

        mark_error_in_tree(files_tree, f)  # 💥 ВОТ ЭТО
def set_mode(mode):
    for frame in (center_frame, create_frame, delete_frame):
        frame.pack_forget()
    global current_mode
    global active_button

    current_mode = mode

    # 💥 Возвращаем обычный текст без галочки
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

    # 💥 Добавляем галочку к выбранной кнопке
    if mode == "MOVE":
        move_btn.config(relief="flat", bg=THEMES[current_theme]["accent"], fg=THEMES[current_theme]["fg"], text=" ✔ MOVE")
        active_button = move_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "COPY":
        duplicate_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text=" ✔ COPY")
        active_button = duplicate_btn
        center_frame.pack(fill="both", expand=True)

    elif mode == "CREATE":
        create_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text=" ✔ CREATE")
        active_button = create_btn
        create_frame.pack(fill="both", expand=True)

    elif mode == "DELETE":
        delete_btn.config(relief="flat", bg=THEMES[current_theme]["accent"],  fg=THEMES[current_theme]["fg"], text=" ✔ DELETE")
        active_button = delete_btn
        delete_frame.pack(fill="both", expand=True)

    action_btn.config(text=mode)
def clear_other_selection(event, tree):
    if event.state & 0x4:
        return

    if tree != files_tree:
        files_tree.selection_remove(files_tree.selection())

    if tree != dest_tree:
        dest_tree.selection_remove(dest_tree.selection())

def execute_action():

    if current_mode == "MOVE":
        move_files()

    elif current_mode == "COPY":
        duplicate_files()

    elif current_mode == "DELETE":
        delete_files()

    elif current_mode == "CREATE":
        create_files()

#drag&drop
def drop_files(event):
    files = root.tk.splitlist(event.data)

    added = []  # 💥 реально добавленные
    skipped = []  # 💥 дубликаты

    for f in files:
        f = norm(f)

        # ❌ если уже есть — пропускаем
        if f in added_paths:
            skipped.append(f)
            continue

        added_paths.add(f)

        icon = get_file_icon(f)

        files_tree.insert(
            "",
            "end",
            text=os.path.basename(f),
            image=icon if icon else "",
            values=(f,)
        )

        if icon:
            icons_cache.append(icon)

        added.append(f)
    adjust_tree_column_full(files_tree)
    # 💥 сообщения
    if added:
        names = [os.path.basename(f) for f in added]
        msg = format_items(len(added), 0, names if len(added)==1 else None)
        show_status(msg)

    elif skipped:
        show_status("All items already added")

def open_log_console():
    global log_text

    console = tk.Toplevel(root)
    console.title("Console Log")
    console.geometry("600x400")

    log_text = tk.Text(console, bg="black", fg="white")
    log_text.pack(fill="both", expand=True)

    for line in log_buffer:
        log_text.insert("end", line + "\n")

    log_text.see("end")

def log_to_console(message):
    global log_buffer

    print(message)
    log_buffer.append(f"> {message}")

    try:
        if log_text and log_text.winfo_exists():
            log_text.insert("end", message + "\n")
            log_text.see("end")
    except:
        pass



def restore_selected_file():
    selected = delete_bin_tree.selection()
    
    for item in selected:
        path = delete_bin_tree.item(item, "values")[0]
        restore_file(path)

        delete_bin_tree.delete(item)
def clear_other_selection(current_tree):
    if current_tree != files_tree:
        files_tree.selection_remove(files_tree.selection())

    if current_tree != dest_tree:
        dest_tree.selection_remove(dest_tree.selection())

#deleting files in list
def delete_selected_file():
    paths = []

    for item in files_tree.selection():
        all_paths = collect_all_children(files_tree, item)

        for p in all_paths:
            excluded_paths.add(norm(p))
            paths.append(p)

        files_tree.delete(item)

    if paths:
        remove_items(paths)
def collect_all_children(tree, item):
    paths = []

    values = tree.item(item, "values")
    if values:
        paths.append(values[0])

    for child in tree.get_children(item):
        paths.extend(collect_all_children(tree, child))

    return paths
#all files clearing for files and destination list 
def clear_all_files():
    paths = get_all_tree_items(files_tree)

    files = sum(1 for p in paths if os.path.isfile(p))
    folders = sum(1 for p in paths if os.path.isdir(p))

    for item in files_tree.get_children():
        files_tree.delete(item)

    selected_files.clear()
    selected_files_set.clear()
    icons_cache.clear()
    added_paths.clear()
    msg = format_items(files, folders, None, "Removed")
    show_status(msg)
    adjust_tree_column_full(files_tree)
    root.after(5000, lambda: set_status("Ready"))

def insert_folder(tree, parent, folder_path, depth=1, max_depth=float("inf"), visited=None):
    if visited is None:
        visited = set()

    real_path = os.path.realpath(folder_path)
    parts = real_path.lower().split(os.sep)
    name = os.path.basename(folder_path) or folder_path
    icon = get_file_icon(folder_path)

    if real_path in visited:
        print("LOOP DETECTED:", real_path)
        return
    visited.add(real_path)
    if parent:
        parent_text = tree.item(parent, "text").strip()
        if parent_text == name:
            print("REPEATING NAME STOP:", name)
            return


    if parts.count(name.lower()) > 2:
        print("RECURSIVE PATTERN:", real_path)
        return
    if depth > 60:
        print("DEPTH LIMIT:", folder_path)
        return


    node = tree.insert(
        parent,
        "end",
        text=" " + name,
        image=icon if icon else "",
        open=True,
        values=(folder_path,)
    )

    if icon:
        icons_cache.append(icon)

    try:
        with os.scandir(folder_path) as entries:

            print("SCAN:", folder_path)

            entries = list(entries)

            for entry in entries:
                print("ENTRY:", entry.name)
                if entry.is_symlink():
                        continue
                full_path = entry.path  

                if entry.is_dir(follow_symlinks=False):
                    insert_folder(tree, node, full_path, depth + 1, max_depth, visited)
                else:
                    icon = get_file_icon(full_path)

                    tree.insert(
                        node,
                        "end",
                        text=" " + entry.name,
                        image=icon if icon else "",
                        values=(full_path,)
                    )

                    if icon:
                        icons_cache.append(icon)

    except Exception as e:
        print("scan error:", e)
def auto_format_date(event):
    entry = event.widget

    if getattr(entry, "is_placeholder", False):
        return

    text = entry.get()

    if not text or getattr(entry, "is_placeholder", False):
        return

    raw = re.sub(r"\D", "", text)

    fmt = date_format_var.get()

    # 💥 только цифры
    if not raw.isdigit():
        return

    result = None

    # 🇷🇺 DMY → 21012000
    if fmt == "DMY" and len(raw) == 8:
        d, m, y = raw[:2], raw[2:4], raw[4:]
        result = f"{d}.{m}.{y}"

    # 🇺🇸 MDY → 01212000
    elif fmt == "MDY" and len(raw) == 8:
        m, d, y = raw[:2], raw[2:4], raw[4:]
        result = f"{m}/{d}/{y}"

    # 🌍 YMD → 20000121
    elif fmt == "YMD" and len(raw) == 8:
        y, m, d = raw[:4], raw[4:6], raw[6:]
        result = f"{y}-{m}-{d}"

    if result:
        entry.delete(0, tk.END)
        entry.insert(0, result)
    update_preview()
def on_click_toggle(event):
    tree = event.widget
    item = tree.identify_row(event.y)

    if not item:
        return

    selected = tree.selection()

    if item in selected:
        tree.selection_remove(item)
    else:

        if not (event.state & 0x4):
            tree.selection_set(item)
        else:
            tree.selection_add(item)

def choose_source_folder():
    folder = filedialog.askdirectory()

    if not folder:
        return

    folder = norm(folder)

    # 💥 проверка на дубликат
    if folder in added_paths:
        show_status("Folder already added")
        return

    added_paths.add(folder)

    insert_folder(files_tree, "", folder)

    folders, files = count_inside_folder(folder)
    name = os.path.basename(folder)

    if folders + files == 0:
        msg = f"Added folder {name}"
    elif folders + files == 1:
        if files == 1:
            msg = f"Added file {name}"
        else:
            msg = f"Added folder {name}"
    else:
        msg = format_items(files, folders)

    show_status(msg)
    adjust_tree_column_full(files_tree)

def clear_destination():
    global destination_folder
    destination_folder = ""

    for item in dest_tree.get_children():
        dest_tree.delete(item)

    dest_tree.column("#0", width=220)
    show_status("Destination cleared")
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
def clear_selection(event=None):
    files_tree.selection_remove(files_tree.selection())
    dest_tree.selection_remove(dest_tree.selection())
    set_status("Selection cleared")

def start_drag_select(event):
    global drag_select_start
    region = files_tree.identify("region", event.x, event.y)
    element = files_tree.identify_element(event.x, event.y)

    if "indicator" in element:
        return

    is_dragging = True

    tree = event.widget
    item = tree.identify_row(event.y)

    if not item:
        return

    selected = tree.selection()

    # toggle
    if item in selected:
        tree.selection_remove(item)
        drag_select_start = None
        return "break"

    if not (event.state & 0x4):
        clear_all_tree_selections(event, tree)
        tree.selection_set(item)
    else:
        tree.selection_add(item)

    drag_select_start = item
    clear_all_tree_selections(event, tree)


def clear_all_tree_selections(event, current_tree):
    global is_dragging

    if event.state & 0x4 or is_dragging:
        return

    for tree in (files_tree, dest_tree, delete_tree):
        if tree != current_tree:
            tree.selection_remove(tree.selection())

def drag_select(event):
    global drag_select_start

    tree = event.widget

    if not drag_select_start:
        return

    current = tree.identify_row(event.y)
    if not current:
        return

    def get_all_items(parent=""):
        result = []
        for child in tree.get_children(parent):
            result.append(child)
            result.extend(get_all_items(child))
        return result

    items = get_all_items()

    try:
        start_index = items.index(drag_select_start)
        current_index = items.index(current)
    except ValueError:
        return

    tree.selection_remove(tree.selection())

    low = min(start_index, current_index)
    high = max(start_index, current_index)

    for i in range(low, high + 1):
        tree.selection_add(items[i])
def stop_drag_select(event):
    global drag_select_start, is_dragging
    drag_select_start = None
    is_dragging = False
#icon fixup
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def adjust_tree_column(tree, name):
    font = tkfont.nametofont("TkDefaultFont")

    name_width = font.measure(name) + 20

    tree.column("#0", width=max(tree.column("#0", "width"), name_width))

def adjust_delete_column_full():
    items = delete_tree.get_children()

    if not items:
        delete_tree.column("#0", width=100, stretch=False)
        return
    
    font = tkfont.nametofont("TkDefaultFont")

    max_width = 220

    for item in delete_tree.get_children():
        text = delete_tree.item(item, "text")
        width = font.measure(text) + 40

        if width > max_width:
            max_width = width

    MAX_LIMIT = 600
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
class TreeTooltip:
    def __init__(self, tree):
        self.tree = tree
        self.tip = None
        self.current_item = None
        self.timer = None

        self.check_mouse()

    def check_mouse(self):
        try:
            x = self.tree.winfo_pointerx()
            y = self.tree.winfo_pointery()

            # переводим в координаты tree
            rel_x = x - self.tree.winfo_rootx()
            rel_y = y - self.tree.winfo_rooty()

            item = self.tree.identify_row(rel_y)

            if item != self.current_item:
                self.cancel()
                self.current_item = item

                if item:
                    self.timer = self.tree.after(800, self.show_tooltip)

        except:
            pass

        self.tree.after(150, self.check_mouse)

    def show_tooltip(self):
        if not self.current_item:
            return

        values = self.tree.item(self.current_item, "values")
        if not values:
            return

        text = values[0]

        x = self.tree.winfo_pointerx() + 15
        y = self.tree.winfo_pointery() + 15

        self.tip = tw = tk.Toplevel(self.tree)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw,
            text=text,
            background="#1e1e1e",
            foreground="white",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9),
            padx=6,
            pady=3
        )
        label.pack()

    def cancel(self):
        if self.timer:
            self.tree.after_cancel(self.timer)
            self.timer = None

        if self.tip:
            self.tip.destroy()
            self.tip = None

        self.current_item = None
def show_file_menu(event):
    global drag_select_start, is_dragging

    tree = files_tree
    item = tree.identify_row(event.y)

  
    drag_select_start = None
    is_dragging = False

    if not item:
        return

    
    if item not in tree.selection():
        tree.selection_set(item)

    file_menu.post(event.x_root, event.y_root)

def get_selected_delete_path():
    selection = delete_tree.selection()
    if not selection:
        return None

    index = delete_tree.index(selection[0])

    if index < len(delete_files_list):
        return delete_files_list[index]

    return None
def reveal_in_explorer(path):
    path = os.path.abspath(path)

    print("RAW PATH:", path)

    folder = os.path.dirname(path)
    base = os.path.splitext(os.path.basename(path))[0]

    if not os.path.exists(folder):
        print("folder missing")
        return

    real_path = None

    for f in os.listdir(folder):
        if os.path.splitext(f)[0] == base:
            real_path = os.path.join(folder, f)
            break

    if not real_path:
        print("file not found, opening folder")
        subprocess.run(f'explorer "{folder}"')
        return

    print("REAL FILE:", real_path)

    subprocess.run(f'explorer /select,"{real_path}"')
    print("REAL PATH:", repr(path))

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
    global drag_select_start, is_dragging

    tree = delete_tree
    item = tree.identify_row(event.y)

    drag_select_start = None
    is_dragging = False

    if not item:
        return

    if item not in tree.selection():
        tree.selection_set(item)

    delete_menu.post(event.x_root, event.y_root)

def clear_highlight(event):

    for i in dest_tree.get_children():
        dest_tree.item(i, tags=())

def highlight_file(event):
    region = files_tree.identify("region", event.x, event.y)
    item = files_tree.identify_row(event.y)

    clear_all_tags(files_tree)

    if item and region == "tree":
        files_tree.item(item, tags=("highlight",))
        files_tree.config(cursor="hand2")
    else:
        files_tree.config(cursor="")

def clear_all_tags(tree, item=""):
    for child in tree.get_children(item):
        tree.item(child, tags=())
        clear_all_tags(tree, child)

def clear_file_highlight(event):
    clear_all_tags(files_tree)

def delete_file_key(event=None):
    selected = list(files_tree.selection())
    paths = []

    for item in selected:
        if not files_tree.exists(item):
            continue

        values = files_tree.item(item, "values")
        if values:
            paths.append(values[0])

        files_tree.delete(item)

    if paths:
        remove_items(paths)

def delete_destination_key(event):
    clear_destination()


def select_all_files(event=None):

    files_tree.selection_set(files_tree.get_children())

    return "break"
def select_all(event):
    event.widget.selection_set(event.widget.get_children())
    return "break"
def on_safe_toggle():
    if not safe_mode.get():
        result = messagebox.askyesno(
            "WARNING",
            "All files you chosen for delete - will be permanently deleted!\nContinue?"
        )
        if not result:
            safe_mode.set(True)
def switch_create_mode(mode):

    preview_mode_btn.config(relief="raised")
    info_mode_btn.config(relief="raised")

    if mode == "preview":
        preview_mode_btn.config(relief="sunken")
        mode_label.config(text="Preview of files to be created")

        update_preview()


        preview_frame.grid(row=0, column=0, sticky="nsew")


        content_box.grid_remove()
        content_scroll.grid_remove()
        content_scroll_x.grid_remove()


        preview_scroll_x.grid()
        preview_scroll_y.grid()

    elif mode == "info":
        info_mode_btn.config(relief="sunken")
        mode_label.config(text="Content of the generated files")


        preview_frame.grid_remove()
        preview_scroll_x.grid_remove()
        preview_scroll_y.grid_remove()


        content_box.grid(row=0, column=0, sticky="nsew")
        content_scroll.grid(row=0, column=1, sticky="ns")
        content_scroll_x.grid(row=1, column=0, sticky="ew")

def show_about():
    global about_window

    if about_window and about_window.winfo_exists():
        about_window.lift()
        return

    about_window = tk.Toplevel(root)
    about_window.title("About PyFiles")
    about_window.geometry("350x350")
    about_window.config(bg="#1e1e1e")

    def on_close():
        global about_window
        about_window.destroy()
        about_window = None

    about_window.protocol("WM_DELETE_WINDOW", on_close)

    # 💥 ИКОНКА ОКНА
    about_window.iconbitmap(resource_path("pyfiles2.ico"))

    # 💥 ЛОГО + TITLE
    logo_img = tk.PhotoImage(file=resource_path("pyfilespng.png"))
    logo_img = logo_img.subsample(16, 16)
    about_window.logo_img = logo_img

    title_label = tk.Label(
        about_window,
        image=logo_img,
        text=" PyFiles v0.11.1",
        compound="left",
        font=("Segoe UI", 12, "bold"),
        bg="#1e1e1e",
        fg="white"
    )
    title_label.pack(pady=10)

    # 💥 ТЕКСТ
    info_label = tk.Label(
        about_window,
        text=(
            "What's that?\n"
            "Not-a-Simple GUI tool\n"
            "\n"
            "Built on:\n"
            "Python 3.14 + tkinter\n"
            "\n"
            "Current version:\n"
            "v0.11.1 - UNDER CONSTRUCTION"
        ),
        bg="#1e1e1e",
        fg="white",
        font=("Segoe UI", 10),
        justify="left"
    )
    info_label.pack(pady=5)

    # 💥 STEAM КНОПКА
    steam_frame = tk.Frame(about_window, bg="#1e1e1e")
    steam_frame.pack(pady=15)

    steam_img = tk.PhotoImage(file=resource_path("steamicon.png"))
    steam_img = steam_img.subsample(30, 30)
    about_window.steam_img = steam_img

    steam_btn = tk.Button(
        steam_frame,
        image=steam_img,
        text="  steam: id/applejuicy23",
        compound="left",
        bg="#2a2a2a",
        fg="cyan",
        activebackground="#3a3a3a",
        activeforeground="white",
        relief="flat",
        bd=1,
        cursor="hand2",
        command=lambda: webbrowser.open("https://steamcommunity.com/id/applejuicy23")
    )
    steam_btn.pack()

    # 💥 HOVER ЭФФЕКТ
    def on_enter(e):
        steam_btn.config(bg="#3a3a3a")

    def on_leave(e):
        steam_btn.config(bg="#2a2a2a")

    steam_btn.bind("<Enter>", on_enter)
    steam_btn.bind("<Leave>", on_leave)

def show_whats_new():
    global whatsnew_window

    if whatsnew_window and whatsnew_window.winfo_exists():
        whatsnew_window.lift()
        return


    whatsnew_window = tk.Toplevel(root)
    whatsnew_window.title("What's New?")
    whatsnew_window.geometry("750x450")
    whatsnew_window.config(bg="#1e1e1e")
    logo_img = tk.PhotoImage(file=resource_path("pyfilespng.png"))
    logo_img = logo_img.subsample(16, 16)
    whatsnew_window.logo_img = logo_img  # чтобы не исчезло
    def on_close():
        global whatsnew_window
        whatsnew_window.destroy()
        whatsnew_window = None

    whatsnew_window.protocol("WM_DELETE_WINDOW", on_close)

    title_label = tk.Label(
        whatsnew_window,
        image=logo_img,
        text=" PyFiles v0.11.1",
        compound="left",
        font=("Segoe UI", 12, "bold"),
        bg="#1e1e1e",
        fg="white"
    )
    title_label.pack(pady=10)

    text = tk.Text(
        whatsnew_window,
        bg="#1e1e1e",
        fg="white",
        bd=0,
        highlightthickness=0
    )
    text.pack(expand=True, fill="both", padx=10, pady=10)

    text.insert("1.0",
        "\n"
        "MAJOR UPDATES: \n"
        "\n"
        ">> Sidebar UI is updated: Move, Copy, Create, and Delete buttons now have beautiful icons along with dynamic checkmarks!\n"
        ">> Cache function inside folders is completely fixed: removing specific files from a folder in the list now correctly sends them to the cache bin instead of moving them to the destination\n"
        ">> Drag & Drop function is fixed (UAC restriction logic adjusted)\n"
        "\n"
        "OTHER UPDATES & FIXES: \n"
        "\n"
        ">> Fixed application crash when holding down the Delete button\n"
        ">> Fixed pathlib error in undo_delete function\n"
        ">> Fixed critical crash in Copy mode when an error occurs before files are processed\n"
        ">> Fixed a bug with icons loading order on application startup\n"
        ">> Sidebar buttons sizing logic rewritten to pixel measurements to properly support images\n"

    )

    text.config(state="disabled")

    whatsnew_window.iconbitmap(resource_path("pyfiles2.ico"))
    whatsnew_window.title("What's New?")
def refresh_all():
    try:
        search_files()
        load_recycle_bin()
        set_status("Refreshed List & Bin")
        log_to_console("Refreshed List & Bin")
    except Exception as e:
        show_error(str(e))
def switch_delete_tab(tab):
    
    current_delete_tab.set(tab)

    delete_tree.grid_forget()
    delete_scroll_y.grid_forget()
    delete_scroll_x.grid_forget()

    delete_bin_tree.pack_forget()
    current_delete_tab.set(tab)

    delete_tree.grid_remove()
    delete_bin_tree.grid_remove()

    if tab == "list":
        delete_tree.grid()

    elif tab == "bin":
        delete_bin_tree.grid()
        load_recycle_bin()
    
    if tab == "list":
        list_btn.config(relief="sunken")
        bin_btn.config(relief="raised")
    else:
        list_btn.config(relief="raised")
        bin_btn.config(relief="sunken")

    delete_tree.grid_forget()
    delete_scroll_y.grid_forget()
    delete_scroll_x.grid_forget()
    delete_bin_tree.pack_forget()

    if tab == "list":
        delete_label.config(text="List of the files on delete")

        delete_tree.grid(row=0, column=0, sticky="nsew")
        delete_scroll_y.grid(row=0, column=1, sticky="ns")
        delete_scroll_x.grid(row=1, column=0, sticky="ew")
    if tab == "bin":
        load_recycle_bin()
        delete_bin_tree.grid(row=0, column=0, sticky="nsew")
        bin_scroll_y.grid(row=0, column=1, sticky="ns")
        bin_scroll_x.grid(row=1, column=0, sticky="ew")

    else:
        delete_label.config(text="Recycle bin for deleted files")
def format_time():
    custom = get_entry_value(time_entry)
    fmt = time_format_var.get()

    dt = None

    if custom:
        try:
            if fmt == "24":
                dt = datetime.strptime(custom, "%H:%M:%S")
            else:
                dt = datetime.strptime(custom, "%I:%M:%S")
        except:
            dt = None

    if not dt:
        dt = datetime.now()

    if fmt == "24":
        return dt.strftime("%H:%M:%S")
    else:
        return dt.strftime("%I:%M:%S")
def apply_prefix(prefix, name, num, ext):
    result = prefix

    result = result.replace("{name}", name)
    result = result.replace("{file}", name)

    if "{num}" in result:
        result = result.replace("{num}", str(num))

    if "{date}" in result:
        result = result.replace("{date}", format_date())

    if "{time}" in result:
        result = result.replace("{time}", format_time())

    return result
def load_bin():
    delete_bin_tree.delete(*delete_bin_tree.get_children())
    bin_icons_cache.clear()

    meta = load_meta()

    for file_id, data in meta.items():

        if os.path.exists(data["original_path"]):
            icon = get_file_icon(data["original_path"])
        else:
            icon = get_file_icon(data["name"])

        delete_bin_tree.insert(
            "",
            "end",
            text=data["name"],
            values=(file_id,),
            image=icon if icon else ""
        )

        if icon:
            bin_icons_cache.append(icon)

def delete_files(all=False):
    paths = []
    for i in delete_tree.selection():
        if not delete_tree.exists(i):
            continue

        values = delete_tree.item(i, "values")
        if not values:
            continue

        paths.append(values[0])

    if not paths:
        show_error("Nothing to delete")
        return

    logs = []
    errors = []
    files = 0
    folders = 0
    names = []

    start_console_log("DELETING FILES", len(paths))

    progress["maximum"] = len(paths)
    progress["value"] = 0

    for i, path in enumerate(paths, start=1):
        try:
            path = str(Path(path).resolve())
            filename = os.path.basename(path)
            if os.path.isfile(path):
                files += 1
            elif os.path.isdir(path):
                folders += 1

            names.append(filename)

            if safe_mode.get():
               
                meta = load_meta()
                file_id = str(uuid.uuid4())
                dst = os.path.join(BIN_DIR, file_id)

                shutil.move(path, dst)

                meta[file_id] = {
                    "original_path": path,
                    "name": filename
                }

                save_meta(meta)

                msg = f"[BIN] {filename} moved to bin"

            else:
             
                if os.path.isfile(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)

                msg = f"[DEL] {filename} deleted"

            logs.append(msg)
            log_to_console(msg)

        except Exception as e:
            err = f"{os.path.basename(path)} >> ERROR: {str(e)}"
            errors.append(err)
            logs.append(err)
            log_to_console(err)

        animate_progress(i)
        root.update()


    save_log(logs, errors)
    if safe_mode.get():
        action = "sent to bin"
    else:
        action = "Deleted"

    # 💥 одиночный случай
    if len(names) == 1:
        if files == 1:
            msg = f"File {names[0]} {action}"
        else:
            msg = f"Folder {names[0]} {action}"
    else:
        msg = format_items(files, folders, None, action.capitalize())

    show_status(msg)
    root.after(5000, reset_progress)


    root.after(50, search_files)
    root.after(50, load_bin)
def load_meta():
    with open(META_FILE, "r") as f:
        return json.load(f)

def save_meta(meta):
    with open(META_FILE, "w") as f:
        json.dump(meta, f, indent=4)

def copy_paths():
    selected = files_tree.selection()

    if not selected:
        return

    paths = []

    for item in selected:
        values = files_tree.item(item, "values")
        if not values:
            continue

        path = values[0]
        paths.append(f'"{path}"')

    result = "\n".join(paths)

    root.clipboard_clear()
    root.clipboard_append(result)

import pythoncom
import win32com.client

def delete_from_bin():
    selected = delete_bin_tree.selection()
    meta = load_meta()

    for item in selected:
        file_id = delete_bin_tree.item(item, "values")[0]
        data = meta.get(file_id)

        path = os.path.join(BIN_DIR, file_id)
        name = data["name"] if data else file_id

        try:
            if os.path.exists(path):
                os.remove(path)

                msg = f"[BIN DEL] {name} permanently deleted"
                log_to_console(msg)

        except Exception as e:
            err = f"{name} >> ERROR: {str(e)}"
            log_to_console(err)
            continue

        meta.pop(file_id, None)

    save_meta(meta)
    load_bin()


import os
import shutil

def restore_file():
    selected = delete_bin_tree.selection()
    meta = load_meta()

    for item in selected:
        file_id = delete_bin_tree.item(item, "values")[0]
        data = meta.get(file_id)

        if not data:
            continue

        src = os.path.join(BIN_DIR, file_id)
        dst = data["original_path"]

        try:
            shutil.move(src, dst)
        except Exception as e:
            print("restore error:", e)
            continue

        meta.pop(file_id)

    files = len(selected)
    msg = f"Restored {files} file{'s' if files != 1 else ''}"
    show_status(msg)


    save_meta(meta)
    load_bin()
    search_files()

def start_console_log(title, count):
    log_to_console(f"---{title} ({count})---")
def set_operation_status(action, logs, errors):
    done = len(logs) - len(errors)
    total = len(logs)

    set_status(f"{action}! Done: {done} | Errors: {len(errors)} | Total: {total} items")


def undo_delete(event=None):
    if not deleted_stack:
        return

    item = deleted_stack.pop()

    name = item["name"]
    folder = item["folder"]

    recycle_bin = Path(os.environ["USERPROFILE"]) / "$Recycle.Bin"

    for root, dirs, files in os.walk(recycle_bin):
        if name in files:
            src = os.path.join(root, name)
            dst = os.path.join(folder, name)

            try:
                shutil.move(src, dst)
                set_success(f"Restored: {name}")
            except Exception as e:
                show_error(str(e))
            return

    show_error("File not found in recycle bin")

def create_files():
    folder = dest_entry.get().strip()

    if not folder or not os.path.isdir(folder):
        show_error("Select valid folder")
        return

    name = get_entry_value(file_name_entry)
    ext = get_entry_value(format_entry)
    prefix = get_prefix_value()
    content = content_box.get("1.0", "end").strip()

    if not name:
        name = "file"

    numbers = generate_numbers()

    if not numbers:
        show_error("Invalid count")
        return

    try:
        progress["maximum"] = len(numbers)
        progress["value"] = 0

        logs = []
        errors = []

        start_console_log("CREATING FILES", len(numbers))

        for index, i in enumerate(numbers, start=1):
            try:
                new_name = apply_prefix(prefix, name, i, ext)
                filename = sanitize_filename(f"{new_name}{ext}")
                path = os.path.join(folder, filename)

                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)

                msg = f"Created {filename} ({folder})"
                logs.append(msg)
                log_to_console(msg)

            except Exception as e:
                err = f"{filename} >> ERROR: {str(e)}"
                errors.append(err)
                logs.append(err)
                log_to_console(err)

            animate_progress(index)
            root.update()

        save_log(logs, errors)

        files = len(logs) - len(errors)

        msg = f"Created {files} file{'s' if files != 1 else ''}"
        show_status(msg)

        root.after(5000, reset_progress)

    except Exception as e:
        show_error(str(e))

def adjust_tree_column_full(tree):
    font = tkfont.nametofont("TkDefaultFont")

    max_width = 0

    def check(item, level=0):
        values = tree.item(item, "values")
        text = tree.item(item, "text")

        # 💥 если есть полный путь — используем его
        if values:
            display_text = values[0]
        else:
            display_text = text

        indent = level * 25

        width = font.measure(display_text) + indent + 80

        nonlocal max_width
        max_width = max(max_width, width)

        for child in tree.get_children(item):
            check(child, level + 1)

    for item in tree.get_children(""):
        check(item)

    tree.column("#0", width=max_width)
def show_create_preview():

    preview = tk.Toplevel(root)
    preview.title("Preview new files")
    preview.geometry("450x350")

    preview_tree = ttk.Treeview(preview)
    preview_tree.pack(fill="both", expand=True)

    preview_tree.tag_configure("new", foreground="blue")


    preview.icons_cache = []

    name = get_entry_value(file_name_entry)
    ext = get_entry_value(format_entry)
    prefix = prefix_entry.get()

    count = 10


    icon = get_create_preview_icon(ext)

    for i in range(count):

        new_name = apply_prefix(prefix, name, i, ext)
        filename = f"{new_name}{ext}"

        preview_tree.insert(
            "",
            "end",
            text=filename,
            image=icon if icon else "",
            tags=("new",)
        )

        if icon:
            preview.icons_cache.append(icon)
            
def get_create_preview_icon(ext):
    if not ext:
        return unknown_icon  # 💥 ВОТ ЭТО ГЛАВНОЕ

    if not ext.startswith("."):
        ext = "." + ext

    ext = ext.lower()

    if ext in IMAGE_EXTENSIONS:
        return image_icon

    hicon = get_icon_by_extension(ext)

    if not hicon:
        return None

    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, 16, 16)

    hdc_mem = hdc.CreateCompatibleDC()
    hdc_mem.SelectObject(hbmp)
    
    win32gui.DrawIconEx(
        hdc_mem.GetHandleOutput(),
        0, 0,
        hicon, 16, 16,
        0, None,
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
    hdc_mem.DeleteDC()
    hdc.DeleteDC()

    return ImageTk.PhotoImage(img)
def show_bin_menu(event):
    item = delete_bin_tree.identify_row(event.y)
    if not item:
        return

    if item not in delete_bin_tree.selection():
        delete_bin_tree.selection_set(item)

    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="Restore", command=restore_file)
    menu.add_command(label="Delete", command=delete_from_bin)

    menu.post(event.x_root, event.y_root)


def restore_selected_from_item(item):
    filename = delete_bin_tree.item(item, "text")

    restore_file(filename)

    delete_bin_tree.delete(item)
def normalize_path(p):
    return os.path.normcase(os.path.normpath(p))


def load_recycle_bin():
    load_bin()

def get_all_tree_items(tree):
    items = []

    def collect(item):
        values = tree.item(item, "values")
        if values:
            items.append(values[0])

        for child in tree.get_children(item):
            collect(child)

    for item in tree.get_children():
        collect(item)

    return items
def bind_auto_update(widget):
    widget.bind("<KeyRelease>", lambda e: (update_create_ui(), update_preview()))
def show_create():
    main_frame.pack(fill="both", expand=True)
    create_frame.pack(fill="both", expand=True)
def delete_file_key_real(event=None):
    selected = delete_tree.selection()

    if not selected:
        show_error("Nothing to delete")
        return

    for item in selected:
        values = delete_tree.item(item, "values")

        if not values:
            continue

        path = values[0]

        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)

            delete_tree.delete(item)

        except Exception as e:
            show_error(str(e))
def sanitize_filename(name):
    forbidden = r'<>:"/\\|?*'
    for char in forbidden:
        name = name.replace(char, "_")
    return name
def parse_count_input(text):
    text = text.strip()

    if not text:
        return [1]

    result = []

    parts = text.split(",")

    for part in parts:
        part = part.strip()

        if ":" in part:
            try:
                start, end = map(int, part.split(":"))
                result.extend(range(start, end + 1))
            except:
                continue

        else:
            try:
                result.append(int(part))
            except:
                continue

    if len(result) == 1 and ":" not in text and "," not in text:
        n = result[0]
        return list(range(1, n + 1))

    return result
def generate_numbers():
    prefix = get_prefix_value()

    if "{num}" not in prefix:
        return [1]

    raw = get_entry_value(count_entry)

    if not raw:
        return list(range(1, 11))

    nums = parse_count_input(raw)

    if not nums:
        nums = [1]

    if random_var.get():
        random.shuffle(nums)

    return nums
def format_date():
    custom = get_entry_value(date_entry)
    fmt = date_format_var.get()

    dt = None

    if custom:
        try:
            if fmt == "DMY":
                dt = datetime.strptime(custom, "%d.%m.%Y")
            elif fmt == "MDY":
                dt = datetime.strptime(custom, "%m/%d/%Y")
            elif fmt == "YMD":
                dt = datetime.strptime(custom, "%Y-%m-%d")
        except:
            dt = None

    if not dt:
        dt = datetime.now()

    if fmt == "DMY":
        return dt.strftime("%d.%m.%Y")
    elif fmt == "MDY":
        return dt.strftime("%m/%d/%Y")
    else:
        return dt.strftime("%Y-%m-%d")
def update_date_placeholder():
    fmt = date_format_var.get()

    if fmt == "DMY":
        ph = "DD.MM.YYYY"
    elif fmt == "MDY":
        ph = "MM/DD/YYYY"
    else:
        ph = "YYYY-MM-DD"

    # 💥 если пользователь уже печатает — НЕ ТРОГАЕМ
    if not getattr(date_entry, "is_placeholder", False):
        return

    add_placeholder(date_entry, ph)
def scan_files(base_path, max_depth):
    result = []

    base_path = os.path.normpath(base_path)
    base_depth = base_path.count(os.sep)

    for root, dirs, files in os.walk(base_path):
        current_depth = root.count(os.sep) - base_depth

        if current_depth >= max_depth:
            dirs[:] = []

        for file in files:
            full_path = os.path.join(root, file)
            result.append(full_path)

    return result
def search_files(event=None):
    folder = delete_dest_entry.get().strip()
    keyword = get_entry_value(delete_keyword_entry)
    ext = get_entry_value(delete_ext_entry)

    delete_tree.delete(*delete_tree.get_children())
    delete_files_list.clear()
    icons_cache.clear()

    if not folder or not os.path.isdir(folder):
        return

    try:
        try:
            max_depth = max(0, int(deep_entry.get()))
        except:
            max_depth = 1

        base_depth = folder.count(os.sep)
        folders_map = {folder: ""}

        for root_dir, dirs, files in os.walk(folder):
            current_depth = root_dir.count(os.sep) - base_depth

            if current_depth >= max_depth:
                dirs[:] = []

       
            if root_dir not in folders_map:
                rel_path = os.path.relpath(root_dir, folder)
                parts = rel_path.split(os.sep)

                parent = ""
                current_path = folder

                for part in parts:
                    current_path = os.path.join(current_path, part)

                    if current_path not in folders_map:
                        node = delete_tree.insert(
                            parent,
                            "end",
                            text=part,
                            open=True
                        )
                        folders_map[current_path] = node

                    parent = folders_map[current_path]

            parent_id = folders_map.get(root_dir, "")

          
            for file in sorted(files, key=natural_sort_key):

                if keyword and keyword.lower() not in file.lower():
                    continue

                if ext and ext != ".":
                    if not file.lower().endswith(ext.lower()):
                        continue

                full_path = os.path.join(root_dir, file)

           
                icon = get_file_icon(full_path)

                delete_tree.insert(
                    parent_id,
                    "end",
                    text=file,
                    values=(full_path,),
                    image=icon if icon else "",
                    tags=("delete",)
                )

                if icon:
                    icons_cache.append(icon)

                delete_files_list.append(full_path)

    except Exception as e:
        print("search error:", e)

    delete_tree.update_idletasks()
    adjust_tree_column_full(delete_tree)
    update_delete_scroll()

def refresh_delete_tree():
    folder = delete_dest_entry.get().strip()

    if not folder or not os.path.exists(folder):
        return

    delete_tree.delete(*delete_tree.get_children())

    try:
        max_depth = int(deep_entry.get())
    except:
        max_depth = 1

    insert_folder(delete_tree, "", folder, 1, max_depth)


def save_log(logs, errors, action="Process"):
    from datetime import datetime
    import os

    now = datetime.now()

    date_str = now.strftime("%m/%d/%Y")
    time_str = now.strftime("%H:%M:%S")

    log_dir = os.path.join(os.getcwd(), "LOGS")

    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_filename = os.path.join(
        log_dir,
        f"log_{now.strftime('%Y%m%d_%H%M%S')}.txt"
    )

    with open(log_filename, "w", encoding="utf-8") as f:
        f.write("LOG\n")
        f.write("Version: PyFiles v0.11.1\n")
        f.write(f"Date {date_str}\n")
        f.write(f"Time {time_str}\n\n")

        f.write("==FILE PROCESS==\n")

        for entry in logs:
            f.write(entry + "\n")

        if errors:
            f.write("\n==ERRORS==\n")
            for err in errors:
                f.write(err + "\n")


def get_all_children(tree, item):
    result = []
    children = tree.get_children(item)

    for child in children:
        result.append(child)
        result.extend(get_all_children(tree, child))

    return result

def add_placeholder(entry, placeholder):

    entry.placeholder = placeholder
    entry.is_placeholder = True

    def show_placeholder():
        entry.delete(0, tk.END)
        entry.insert(0, placeholder)
        entry.config(fg="#888888")
        entry.is_placeholder = True

    def hide_placeholder():
        if entry.is_placeholder:
            entry.delete(0, tk.END)
            entry.config(fg=THEMES[current_theme]["fg"])
            entry.is_placeholder = False

    def on_focus_in(event):
        hide_placeholder()

    def on_focus_out(event):
        if not entry.get():
            show_placeholder()

    # 💥 ЧИСТИМ старые бинды
    entry.unbind("<FocusIn>")
    entry.unbind("<FocusOut>")

    # 💥 ТОЛЬКО focus, НИКАКИХ Key
    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

    show_placeholder()

def on_click_toggle(event):
    tree = event.widget
    item = tree.identify_row(event.y)

    if not item:
        return

    selected = tree.selection()

    if item in selected:
        tree.selection_remove(item)
    else:
        if not (event.state & 0x4):
            tree.selection_set(item)
        else:
            tree.selection_add(item)

def update_preview():
    

    preview_tree.delete(*preview_tree.get_children())
    preview_tree.icons_cache = []
    name = get_entry_value(file_name_entry)
    ext = get_entry_value(format_entry)


    icon = get_create_preview_icon(ext)
    prefix = get_prefix_value()

    if not name:
        name = "file"

    numbers = generate_numbers()

    font = tkfont.nametofont("TkDefaultFont")
    max_width = 220
    
    for i in numbers:
        new_name = apply_prefix(prefix, name, i, ext)
        filename = f"{new_name}{ext}"

        preview_tree.insert(
            "",
            "end",
            text=filename,
            image=icon if icon else "",
            tags=("new",)
        )

        width = font.measure(filename) + 40
        if width > max_width:
            max_width = width
        if icon:
            preview_tree.icons_cache.append(icon)

    preview_tree.column("#0", width=max_width, stretch=False)
    preview_tree.update_idletasks()

def get_entry_value(entry):
    val = entry.get().strip()

    # 💥 если это placeholder — игнорим
    if getattr(entry, "is_placeholder", False):
        return ""

    return val
def natural_sort_key(s):
    return [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r'(\d+)', s)
    ]
def update_create_ui():
    prefix = get_entry_value(prefix_entry)

    if "{num}" in prefix:
        count_frame.pack(anchor="w", pady=5)
    else:
        count_frame.pack_forget()

    if "{time}" in prefix:
        time_frame.pack(anchor="w", pady=5)
    else:
        time_frame.pack_forget()

    if "{date}" in prefix:
        date_frame.pack(anchor="w", pady=5)
    else:
        date_frame.pack_forget()
def get_prefix_value():
    prefix = get_entry_value(prefix_entry)

    if not prefix:
        return "{file}"

    return prefix
#if error appeared
def show_error(msg):
    show_error_status(msg)

def set_status(msg):
    status_label.config(text=msg, fg=THEMES[current_theme]["fg"])

def set_success(msg):
    status_label.config(text=msg, fg="darkgreen")


#better to sort it but too much shit, later...
#window
root = TkinterDnD.Tk()
root.title("PyFiles v0.11.1")
root.geometry("1200x800")

IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".webp", ".avif",
    ".bmp", ".gif", ".tiff", ".ico"
}
image_icon = ImageTk.PhotoImage(Image.open(resource_path("photoico.png")).resize((16, 16)))
unknown_icon = ImageTk.PhotoImage(Image.open(resource_path("unknown.png")).resize((16, 16)))

sidebar_move_icon = ImageTk.PhotoImage(Image.open(resource_path("move.png")).resize((20, 20)))
sidebar_copy_icon = ImageTk.PhotoImage(Image.open(resource_path("copy.png")).resize((20, 20)))
sidebar_create_icon = ImageTk.PhotoImage(Image.open(resource_path("create.png")).resize((20, 20)))
sidebar_delete_icon = ImageTk.PhotoImage(Image.open(resource_path("delete.png")).resize((20, 20)))

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
create_settings = {
    "num_mode": "range",
    "num_values": (1, 10), 
    "num_random": False,

    "date_format": "RU",
    "date_custom": None,

    "time_format": "24",
    "time_custom": None 
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
    if current_theme == "default":
        entry_bg = "#ffffff"
        entry_fg = "#000000"
    elif current_theme == "light":
        entry_bg = "#ffffff"
        entry_fg = "#000000"
    else:
        entry_bg = "#2a2a2a"
        entry_fg = "#ffffff"

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
        delete_ext_entry,
        deep_entry,
    ):
        if current_theme == "default":
            entry.config(
                bg="#ffffff",
                fg="#000000",
                insertbackground="#000000",
                selectbackground="#cce8ff",
                selectforeground="#000000"
            )


            if getattr(entry, "is_placeholder", False):
                entry.config(fg="#888888")

        elif current_theme == "light":
            entry.config(
                bg="#ffffff",
                fg="#000000",
                insertbackground="#000000"
            )

        else:
            entry.config(
                bg="#2a2a2a",
                fg="#ffffff",
                insertbackground="#ffffff"
            )
        entry.config(
        bg="#ffffff" if current_theme == "default" else entry_bg,
        fg="#000000",
        insertbackground="#000000",
        selectbackground="#cce8ff",
        selectforeground="#000000"
    )

        if getattr(entry, "is_placeholder", False):
            entry.config(
                fg="#888888",
                bg=entry_bg
            )
        else:
            entry.config(
                fg=entry_fg,
                bg=entry_bg
            )
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
            theme_light, theme_gray, theme_dark, list_btn, bin_btn):
        btn.config(fg=t["fg"])
    for lbl in (theme_light, theme_gray, theme_dark):
        lbl.config(bg=t["bg"], fg=t["fg"])



    left_panel.config(bg=t["panel"])
    bottom_frame.config(bg=t["panel"])


    for btn in (move_btn, duplicate_btn, create_btn, delete_btn, open_console_btn, list_btn, bin_btn):
        btn.config(bg=t["panel"], fg=t["fg"], activebackground=t["accent"])


    about_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]  # 💥 ВОТ ОНО
)
    whatsnew_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]
)
    if current_theme == "dark":
        about_tab.config(bg="#e0e0e0", fg="#000000")
        whatsnew_tab.config(bg="#e0e0e0", fg="#000000")
    else:
        about_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]  # 💥 ВОТ ОНО
)
        whatsnew_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]
)

    status_label.config(bg=t["panel"], fg=t["fg"])
    

    for f in (
        center_frame, create_frame, delete_frame,
        left_create, right_create,
        left_delete, right_delete,
        dest_row, delete_row,
        files_buttons_frame, dest_buttons_frame,
        mode_frame, content_container, delete_container,
        delete_top_frame, delete_tab_frame,
        count_frame, time_frame, date_frame
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
    style.configure(
    "TRadiobutton",
    background=t["bg"],
    foreground=t["fg"]
    )
    style.configure("TButton",
    background=t["panel"],
    foreground=t["fg"]
    )
    style.map("TButton",
        background=[("active", t["accent"])],
        foreground=[("!disabled", "#ffffff" if current_theme == "dark" else t["fg"])]
    )
    style.map("TRadiobutton",
    background=[("active", t["bg"])],
    foreground=[("active", t["fg"])]
)
    safe_check.config(
    bg=t["bg"],
    fg=t["fg"],
    activebackground=t["bg"],
    activeforeground=t["fg"],
    selectcolor=t["panel"],
    highlightthickness=0,
    bd=0
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
    for frame in (count_frame, time_frame, date_frame):
        for widget in frame.winfo_children():
            try:

                if isinstance(widget, tk.Entry):
                    if current_theme == "default":
                        widget.config(
                            bg="#ffffff",
                            fg="#000000",
                            insertbackground="#000000"
                        )
                    elif current_theme == "light":
                        widget.config(
                            bg="#ffffff",
                            fg="#000000",
                            insertbackground="#000000"
                        )
                    else:
                        widget.config(
                            bg="#2a2a2a",
                            fg="#ffffff",
                            insertbackground="#ffffff"
                        )
                    continue 


                widget.config(bg=t["bg"], fg=t["fg"])

            except:
                pass

            if isinstance(widget, (tk.Radiobutton, tk.Checkbutton)):
                widget.config(
                    bg=t["bg"],
                    fg=t["fg"],
                    selectcolor="#ffffff" if current_theme == "default" else t["panel"],
                    activebackground=t["bg"],
                    activeforeground=t["fg"]
                )
                
    if getattr(entry, "is_placeholder", False):
        entry.config(fg="#888888")
    else:
        entry.config(fg=entry_fg)
    for widget in left_create.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])

    for widget in left_delete.winfo_children():
        if isinstance(widget, tk.Label):
            widget.config(bg=t["bg"], fg=t["fg"])
    delete_label.config(bg=t["bg"], fg=t["fg"])
    if current_theme == "dark":
        about_tab.config(bg="#2a2a2a", fg="#ffffff")
        whatsnew_tab.config(bg="#2a2a2a", fg="#ffffff")
    else:
        about_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]  # 💥 ВОТ ОНО
)
        whatsnew_tab.config(
    bg=t["panel"],
    fg=t["fg"],
    activebackground=t["panel"]
)

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
    if last_status:
        status_type, msg = last_status

        if status_type == "error":
            status_label.config(text=f"Error: {msg}", fg="red")
        elif status_type == "success":
            status_label.config(text=msg, fg="green")
        else:
            status_label.config(text=msg, fg=t["fg"])
    if current_theme == "default":
        style.configure(
            "TEntry",
            fieldbackground="#ffffff",  
            background="#ffffff"
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

file_menu.add_command(
    label="Reveal in Explorer",
    command=lambda: os.startfile(
        os.path.dirname(
            files_tree.item(files_tree.selection()[0], "values")[0]
        )
    )
)

file_menu.add_command(
    label="Copy path",
    command=copy_paths
)
delete_menu = tk.Menu(root, tearoff=0)

delete_menu.add_command(label="Open", command=lambda: open_delete_file())
delete_menu.add_command(label="Remove", command=lambda: delete_files())
delete_menu.add_separator()
delete_menu.add_command(
    label="Reveal in Explorer",
    command=lambda: reveal_in_explorer(
        delete_tree.item(delete_tree.selection()[0], "values")[0]
    )
)
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
    text=" PyFiles v0.11.1",
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
    "width": 150,    # Ширина в пикселях (эквивалент старым 14 символам)
    "height": 55,    # Высота в пикселях (эквивалент старым 3 строкам)
    "font": ("Segoe UI", 11),
    "bd": 0,
    "anchor": "w",
    "padx": 15       # Слегка отодвигаем иконку от левого края, чтобы было красиво
}

move_btn = tk.Button(left_panel, text="  MOVE", image=sidebar_move_icon, compound="left", command=lambda: set_mode("MOVE"), **btn_style)
move_btn.pack(fill="x", padx=0, pady=0)

duplicate_btn = tk.Button(left_panel, text="  COPY", image=sidebar_copy_icon, compound="left", command=lambda: set_mode("COPY"), **btn_style)
duplicate_btn.pack(fill="x", padx=0, pady=0)

create_btn = tk.Button(left_panel, text="  CREATE", image=sidebar_create_icon, compound="left", command=lambda: set_mode("CREATE"), **btn_style)
create_btn.pack(fill="x", padx=0, pady=0)

delete_btn = tk.Button(left_panel, text="  DELETE", image=sidebar_delete_icon, compound="left", command=lambda: set_mode("DELETE"), **btn_style)
delete_btn.pack(fill="x", padx=0, pady=0)

about_tab = tk.Button(
    left_panel,
    text="About",
    padx=12,
    pady=3,
    bd=1,
    relief="raised",
    cursor="hand2"
)

about_tab.pack(padx=0, pady=(20,0))

whatsnew_tab = tk.Button(
    left_panel,
    text="What's new?",
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

open_console_btn = tk.Button(
    left_panel,
    text="Open Console",
    command=open_log_console,
    bd=1,
    relief="raised",
    cursor="hand2"
)

open_console_btn.pack(pady=(10, 0), padx=5, fill="x")

theme_light.bind("<Button-1>", lambda e: apply_theme("light"))
theme_gray.bind("<Button-1>", lambda e: apply_theme("default"))
theme_dark.bind("<Button-1>", lambda e: apply_theme("dark"))
about_tab.bind("<ButtonRelease-1>", lambda e: show_about())
whatsnew_tab.bind("<ButtonRelease-1>", lambda e: show_whats_new())

main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

def sidebar_hover(event):
    if event.widget != active_button:
        event.widget.config(bg=THEMES[current_theme]["accent"])

def sidebar_leave(event):
    if event.widget != active_button:
        event.widget.config(bg=THEMES[current_theme]["panel"])

for btn in (move_btn, duplicate_btn, create_btn, delete_btn, open_console_btn):
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

TreeTooltip(files_tree)
files_tree.bind("<<TreeviewOpen>>", lambda e: adjust_tree_column_full(files_tree))
files_tree["show"] = "tree"

files_scrollbar = ttk.Scrollbar(files_frame, orient="vertical", command=files_tree.yview)
files_scrollbar_x = ttk.Scrollbar(files_frame, orient="horizontal", command=files_tree.xview)

files_tree.configure(
    yscrollcommand=files_scrollbar.set,
    xscrollcommand=files_scrollbar_x.set
)
files_tree.update_idletasks()
adjust_tree_column_full(files_tree)
files_tree.yview_moveto(0)
files_tree.xview_moveto(0)

files_tree.grid(row=0, column=0, sticky="nsew")
files_scrollbar.grid(row=0, column=1, sticky="ns")
files_scrollbar_x.grid(row=1, column=0, sticky="ew")

files_frame.grid_rowconfigure(0, weight=1)
files_frame.grid_columnconfigure(0, weight=1)

files_tree.column("#0", width=80, minwidth=100, stretch=False)
files_tree.heading("#0", text="")


files_tree.bind("<Double-1>", open_selected_file)
files_tree.bind("<Motion>", highlight_file)
files_tree.bind("<Leave>", clear_file_highlight)
files_tree.bind("<Button-3>", show_file_menu)
files_tree.bind("<Control-a>", select_all_files)
files_tree.bind("<Button-1>", start_drag_select)
files_tree.bind("<B1-Motion>", drag_select)
files_tree.bind("<ButtonRelease-1>", stop_drag_select)
files_tree.tag_configure("highlight", background="#cce8ff")
files_tree.tag_configure("error", foreground="red")
files_tree.drop_target_register(DND_FILES)
files_tree.dnd_bind('<<Drop>>', drop_files)
files_tree.focus_set()

files_buttons_frame = tk.Frame(center_frame)
files_buttons_frame.grid_columnconfigure(0, weight=1)
files_buttons_frame.grid_columnconfigure(1, weight=1)
files_buttons_frame.grid(row=2, column=0, pady=(10,15))

choose_files_btn = ttk.Button(files_buttons_frame, text="Choose Files", command=choose_files)
choose_files_btn.grid(row=0, column=0, padx=5, sticky="e")

choose_folder_src_btn = ttk.Button(
    files_buttons_frame,
    text="Choose Folder",
    command=choose_source_folder
)

choose_folder_src_btn.grid(row=0, column=2, padx=5)

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
add_placeholder(file_name_entry, "Enter the name")

tk.Label(left_create, text="File format:").pack(anchor="w")

format_entry = tk.Entry(left_create, width=10)
format_entry.pack(anchor="w", pady=5)
add_placeholder(format_entry, ".")

tk.Label(left_create, text="Prefix to name").pack(anchor="w")

prefix_entry = tk.Entry(left_create, width=25)
prefix_entry.pack(anchor="w", pady=5)
prefix_entry.pack(anchor="w", pady=5)
add_placeholder(prefix_entry, "Enter the prefixes")

count_frame = tk.Frame(left_create)

count_label = tk.Label(count_frame, text="Count:")
count_entry = tk.Entry(count_frame, width=10)

random_var = tk.BooleanVar()
random_checkbox = tk.Checkbutton(
    count_frame,
    text="Random",
    variable=random_var,
    command=update_preview
)

count_label.pack(anchor="w")
count_entry.pack(anchor="w", pady=2)
random_checkbox.pack(anchor="w")

add_placeholder(count_entry, "1")

time_frame = tk.Frame(left_create)

time_label = tk.Label(time_frame, text="Time:")

time_entry = tk.Entry(time_frame, width=15)
add_placeholder(time_entry, "HH:MM:SS")

time_format_var = tk.StringVar(value="24")

time_24 = tk.Radiobutton(
    time_frame,
    text="24h",
    variable=time_format_var,
    value="24",
    command=update_preview
)

time_12 = tk.Radiobutton(
    time_frame,
    text="12h",
    variable=time_format_var,
    value="12",
    command=update_preview
)

time_label.pack(anchor="w")
time_entry.pack(anchor="w", pady=2)

time_24.pack(anchor="w")
time_12.pack(anchor="w")

date_frame = tk.Frame(left_create)

date_label = tk.Label(date_frame, text="Date:")

date_format_var = tk.StringVar(value="DMY")

date_entry = tk.Entry(date_frame, width=15)
add_placeholder(date_entry, "DD.MM.YYYY")

date_entry.bind("<KeyRelease>", auto_format_date, add="+")

date_dmy = tk.Radiobutton(
    date_frame,
    text="DD.MM.YYYY",
    variable=date_format_var,
    value="DMY",
    command=lambda: (update_date_placeholder(), update_preview())
)

date_mdy = tk.Radiobutton(
    date_frame,
    text="MM/DD/YYYY",
    variable=date_format_var,
    value="MDY",
    command=lambda: (update_date_placeholder(), update_preview())
)

date_ymd = tk.Radiobutton(
    date_frame,
    text="YYYY-MM-DD",
    variable=date_format_var,
    value="YMD",
    command=lambda: (update_date_placeholder(), update_preview())
)

date_label.pack(anchor="w")
date_entry.pack(anchor="w", pady=2)

date_dmy.pack(anchor="w")
date_mdy.pack(anchor="w")
date_ymd.pack(anchor="w")

bind_auto_update(file_name_entry)
bind_auto_update(format_entry)
bind_auto_update(prefix_entry)
bind_auto_update(count_entry)
bind_auto_update(time_entry)


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
prefix_entry.bind("<KeyRelease>", lambda e: (update_create_ui(), update_preview()))
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
mode_label = tk.Label(
    right_create,
    text="Preview of files to be created"
)
mode_label.pack(anchor="w")

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
add_placeholder(delete_keyword_entry, "Enter the keywords")

tk.Label(left_delete, text="File format:").pack(anchor="w")

delete_ext_entry = tk.Entry(left_delete, width=10)
delete_ext_entry.pack(anchor="w", pady=5)
add_placeholder(delete_ext_entry, ".")


tk.Label(left_delete, text="DeepScan level:").pack(anchor="w")

delete_dest_entry.bind("<KeyRelease>", lambda e: refresh_delete_tree())

deep_entry = tk.Entry(left_delete, width=10)
deep_entry.pack(anchor="w", pady=5)

deep_entry.insert(0, "1")

deep_entry.bind("<KeyRelease>", search_files)






safe_mode = tk.BooleanVar(value=True)
safe_check = tk.Checkbutton(
    left_delete,
    text="Safe Mode",
    variable=safe_mode
)
safe_check.pack(anchor="w", pady=5)
safe_check.config(command=on_safe_toggle)













delete_keyword_entry.bind("<KeyRelease>", search_files)
delete_ext_entry.bind("<KeyRelease>", search_files)
delete_dest_entry.bind("<KeyRelease>", search_files)
deep_entry.bind("<FocusOut>", search_files)

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


delete_top_frame = tk.Frame(right_delete)
delete_top_frame.pack(fill="x")

delete_label = tk.Label(
    delete_top_frame,
    text="List of the files on delete"
)
delete_label.pack(anchor="w")


delete_tab_frame = tk.Frame(delete_top_frame)
delete_tab_frame.pack(anchor="w", pady=(0,5))

delete_container = tk.Frame(right_delete)
delete_container.pack(fill="both", expand=True)

delete_container.grid_rowconfigure(0, weight=1)
delete_container.grid_columnconfigure(0, weight=1)

delete_tree = ttk.Treeview(
    delete_container,
    show="tree",
    selectmode="extended"
)
root.bind("<Delete>", lambda e: delete_files() if current_mode == "DELETE" else None)
delete_tree.bind("<Button-3>", show_delete_menu)
delete_tree.bind("<Button-1>", start_drag_select)
delete_tree.bind("<B1-Motion>", drag_select)
delete_tree.bind("<ButtonRelease-1>", stop_drag_select)
delete_tree.bind("<Control-a>", select_all)
delete_tree.bind("<Control-A>", select_all)
delete_tree.bind("<Escape>", clear_delete_selection)
delete_tree.bind("<Control-z>", undo_delete)

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
delete_tree.tag_configure("delete", foreground="red")
delete_tree["show"] = "tree"

list_btn = tk.Button(
    delete_tab_frame,
    text="List",
    width=8,
    relief="sunken",
    command=lambda: switch_delete_tab("list")
)
list_btn.pack(side="left")

bin_btn = tk.Button(
    delete_tab_frame,
    text="Bin",
    width=8,
    command=lambda: switch_delete_tab("bin")
)
bin_btn.pack(side="left")
delete_bin_tree = ttk.Treeview(delete_container)
delete_bin_tree["show"] = "tree"
bin_scroll_y = ttk.Scrollbar(delete_container, orient="vertical", command=delete_bin_tree.yview)
bin_scroll_x = ttk.Scrollbar(delete_container, orient="horizontal", command=delete_bin_tree.xview)

delete_bin_tree.configure(
    yscrollcommand=bin_scroll_y.set,
    xscrollcommand=bin_scroll_x.set
)

bin_scroll_y.grid(row=0, column=1, sticky="ns")
bin_scroll_x.grid(row=1, column=0, sticky="ew")




delete_bin_tree.config(selectmode="extended")

delete_bin_tree.bind("<Button-1>", start_drag_select)
delete_bin_tree.bind("<B1-Motion>", drag_select)
delete_bin_tree.bind("<ButtonRelease-1>", stop_drag_select)
delete_bin_tree.bind("<Button-3>", show_bin_menu)
delete_bin_tree.bind("<Delete>", lambda e: delete_from_bin())
current_delete_tab = tk.StringVar(value="list")
switch_delete_tab("list")


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
dest_tree.bind("<Button-1>", lambda e: clear_all_tree_selections(e, dest_tree))
# buttons
dest_buttons_frame = tk.Frame(center_frame)
dest_buttons_frame.grid(row=2, column=1, pady=(10,15))
dest_buttons_frame.grid_columnconfigure(0, weight=1)
dest_buttons_frame.grid_columnconfigure(1, weight=1)
files_tree.bind("<Delete>", delete_file_key)
dest_tree.bind("<Delete>", delete_destination_key)
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



def global_hotkeys(event):
    if event.state & 0x4:
        if event.keycode == 67:  # C
            event.widget.event_generate("<<Copy>>")
        elif event.keycode == 86:  # V
            event.widget.event_generate("<<Paste>>")
root.bind_all("<KeyPress>", global_hotkeys)



icons_cache.append(image_icon)
set_mode("MOVE")
switch_create_mode("preview")
apply_theme("default")
update_create_ui()
refresh_delete_tree()
root.bind("<Escape>", clear_selection)
root.update_idletasks()
root.after(50, search_files)
root.mainloop()
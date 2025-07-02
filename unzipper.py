import os
import zipfile
import time
from pathlib import Path
from shutil import copy2
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import sys
import platform
import pystray
from PIL import Image, ImageDraw
import winshell
from win32com.client import Dispatch

# Add rarfile import
try:
    import rarfile
except ImportError:
    rarfile = None


def get_base_dir():
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent

CONFIG_FILE = get_base_dir() / "unzipper_config.txt"

def write_config(monitor_folder, dest_folder, delete_after_zip, delete_after_extracted, file_exts, logic_input=None, copy_enabled=None, logic_enabled=None, copy_whole_folder=None):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"monitor_folder={monitor_folder}\n")
        f.write(f"dest_folder={dest_folder}\n")
        f.write(f"delete_after_zip={str(delete_after_zip)}\n")
        f.write(f"delete_after_extracted={str(delete_after_extracted)}\n")
        f.write(f"file_exts={file_exts}\n")
        if logic_input is not None:
            f.write(f"logic_input={logic_input}\n")
        if copy_enabled is not None:
            f.write(f"copy_enabled={str(copy_enabled)}\n")
        if logic_enabled is not None:
            f.write(f"logic_enabled={str(logic_enabled)}\n")
        if copy_whole_folder is not None:
            f.write(f"copy_whole_folder={str(copy_whole_folder)}\n")

def read_config():
    config = {}
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    config[k] = v
    return config

def get_startup_shortcut_path():
    startup_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    exe_name = Path(sys.argv[0]).name
    shortcut_name = "Unzipper.lnk"
    return os.path.join(startup_dir, shortcut_name)

def create_startup_shortcut():
    if platform.system() != "Windows" or winshell is None or Dispatch is None:
        return False
    shortcut_path = get_startup_shortcut_path()
    target = sys.executable if getattr(sys, 'frozen', False) else sys.executable
    script = sys.argv[0]
    if getattr(sys, 'frozen', False):
        # Running as .exe
        target = sys.executable
        script = sys.argv[0]
    else:
        # Running as .py
        target = sys.executable
        script = os.path.abspath(sys.argv[0])
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = script if getattr(sys, 'frozen', False) else target
    shortcut.WorkingDirectory = os.path.dirname(script)
    shortcut.Arguments = "" if getattr(sys, 'frozen', False) else f'"{script}"'
    shortcut.IconLocation = script
    shortcut.save()
    return True

def remove_startup_shortcut():
    shortcut_path = get_startup_shortcut_path()
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        return True
    return False

def is_startup_enabled():
    return os.path.exists(get_startup_shortcut_path())

class ZipExtractorHandler(FileSystemEventHandler):
    def __init__(
        self, download_folder, target_folder, delete_after_zip=False, delete_after_extracted=False,
        file_exts=None, gui_callback=None, copy_enabled=True, logic_input=None, logic_enabled=False, copy_whole_folder=False
    ):
        self.download_folder = Path(download_folder)
        self.processed_files = set()
        self.target_folder = Path(target_folder)
        self.target_folder.mkdir(parents=True, exist_ok=True)
        self.copy_enabled = copy_enabled
        self.logic_input = logic_input
        self.logic_enabled = logic_enabled
        self.copy_whole_folder = copy_whole_folder
        if file_exts:
            cleaned = [ext.strip().lstrip(".").lower() for ext in file_exts.split(',') if ext.strip()]
            if cleaned:
                self.collect_exts = {ext for ext in cleaned}
            else:
                self.collect_exts = None
        else:
            self.collect_exts = None
        self.archive_exts = {'.zip', '.rar'}
        self.gui_callback = gui_callback
        self.delete_after_zip = delete_after_zip
        self.delete_after_extracted = delete_after_extracted

    def log(self, msg):
        if self.gui_callback:
            self.gui_callback(msg)
        else:
            print(msg)

    def _wait_until_file_ready(self, file_path, timeout=10):
        """Wait until the file is not growing in size (download complete)."""
        import time
        last_size = -1
        stable_count = 0
        start = time.time()
        while time.time() - start < timeout:
            try:
                size = os.path.getsize(file_path)
                if size == last_size:
                    stable_count += 1
                    if stable_count >= 2:
                        return True
                else:
                    stable_count = 0
                last_size = size
            except Exception:
                pass
            time.sleep(0.5)
        return False

    def on_created(self, event):
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        ext = file_path.suffix.lower()
        if ext in self.archive_exts and file_path not in self.processed_files:
            if self._wait_until_file_ready(file_path):
                time.sleep(0.5)
                if ext == '.zip':
                    self.extract_zip(file_path)
                elif ext == '.rar':
                    self.extract_rar(file_path)

    def on_moved(self, event):
        # Handle file renames (e.g., .crdownload -> .zip)
        if event.is_directory:
            return
        file_path = Path(event.dest_path)
        ext = file_path.suffix.lower()
        if ext in self.archive_exts and file_path not in self.processed_files:
            if self._wait_until_file_ready(file_path):
                time.sleep(0.5)
                if ext == '.zip':
                    self.extract_zip(file_path)
                elif ext == '.rar':
                    self.extract_rar(file_path)

    def _copy_entire_folder(self, src_folder):
        import shutil
        dest = self.target_folder / Path(src_folder).name
        counter = 1
        orig_dest = dest
        while dest.exists():
            dest = self.target_folder / f"{orig_dest.stem}_{counter}{orig_dest.suffix}"
            counter += 1
        try:
            shutil.copytree(src_folder, dest)
            self.log(f"Copied entire folder: {src_folder} -> {dest}")
            # Always delete extracted folder after copying if option is enabled
            if self.delete_after_extracted and Path(src_folder).exists():
                try:
                    shutil.rmtree(src_folder)
                    self.log(f"Deleted extracted folder after copying: {src_folder}")
                except Exception as e:
                    self.log(f"Failed to delete extracted folder after copying: {src_folder} ({e})")
        except Exception as e:
            self.log(f"Failed to copy entire folder: {src_folder} -> {dest}: {e}")

    def extract_zip(self, zip_path, stop_event=None):
        try:
            if not zip_path.exists():
                return
            self.log(f"Found new ZIP file: {zip_path.name}")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                root_dirs = set()
                root_files = set()
                for member in zip_ref.namelist():
                    parts = member.split('/')
                    if len(parts) == 1 or (len(parts) == 2 and parts[1] == ''):
                        root_files.add(parts[0])
                    else:
                        root_dirs.add(parts[0])
                root_dirs = {d for d in root_dirs if any(f.startswith(d + '/') for f in zip_ref.namelist())}
                extract_to_downloads = False
                if len(root_dirs) == 1 and not root_files:
                    extract_folder = self.download_folder / list(root_dirs)[0]
                    zip_ref.extractall(self.download_folder)
                    extract_to_downloads = True
                else:
                    extract_folder = self.download_folder / zip_path.stem
                    counter = 1
                    original_extract_folder = extract_folder
                    while extract_folder.exists():
                        extract_folder = Path(f"{original_extract_folder}_{counter}")
                        counter += 1
                    zip_ref.extractall(extract_folder)
            if extract_to_downloads:
                self.log(f"Successfully extracted to monitored folder: {extract_folder}")
                search_folder = extract_folder
            else:
                self.log(f"Successfully extracted to: {extract_folder}")
                search_folder = extract_folder
            # Copy whole folder if enabled
            if self.copy_whole_folder:
                self.log("Copying entire extracted folder to destination (option enabled)...")
                self._copy_entire_folder(search_folder)
                # Always try to delete after copying (handled in _copy_entire_folder)
            else:
                self.copy_selected_files(search_folder, stop_event=stop_event)
            self.processed_files.add(zip_path)
            # Delete ZIP if option is enabled
            if self.delete_after_zip:
                try:
                    if zip_path.exists():
                        zip_path.unlink()
                        self.log(f"Deleted ZIP file: {zip_path}")
                except Exception as e:
                    self.log(f"Failed to delete ZIP file: {zip_path} ({e})")
        except zipfile.BadZipFile:
            self.log(f"Error: {zip_path.name} is not a valid ZIP file or is corrupted")
        except PermissionError:
            self.log(f"Error: Permission denied accessing {zip_path.name}")
        except Exception as e:
            self.log(f"Error extracting {zip_path.name}: {str(e)}")

    def extract_rar(self, rar_path, stop_event=None):
        if not rarfile:
            self.log("Error: rarfile module not installed. Install with 'pip install rarfile'.")
            return
        try:
            if not rar_path.exists():
                return
            self.log(f"Found new RAR file: {rar_path.name}")
            try:
                tool_found = False
                try:
                    tool_found = rarfile.tool_setup() is not None
                except AttributeError:
                    try:
                        tool_found = rarfile._get_unrar_tool() is not None
                    except Exception:
                        pass
                if not tool_found:
                    self.log(
                        "Error: No working unrar/rar tool found. Please install 'unrar' or 'rar' and ensure it is in your PATH.\n"
                        "To install unrar for Windows:\n"
                        "1. Download the Windows binary from https://www.rarlab.com/rar_add.htm (look for 'UnRAR for Windows').\n"
                        "2. Extract the downloaded archive.\n"
                        "3. Copy unrar.exe to a folder in your system PATH (e.g., C:\\Windows or C:\\Windows\\System32), or keep it in the same folder as your script.\n"
                        "4. Optionally, add the folder containing unrar.exe to your PATH environment variable for global access.\n"
                        "5. The rarfile Python package will then be able to use it automatically."
                    )
                    return
                with rarfile.RarFile(rar_path, 'r') as rar_ref:
                    root_dirs = set()
                    root_files = set()
                    for member in rar_ref.namelist():
                        parts = member.split('/')
                        if len(parts) == 1 or (len(parts) == 2 and parts[1] == ''):
                            root_files.add(parts[0])
                        else:
                            root_dirs.add(parts[0])
                    root_dirs = {d for d in root_dirs if any(f.startswith(d + '/') for f in rar_ref.namelist())}
                    extract_to_downloads = False
                    if len(root_dirs) == 1 and not root_files:
                        extract_folder = self.download_folder / list(root_dirs)[0]
                        rar_ref.extractall(self.download_folder)
                        extract_to_downloads = True
                    else:
                        extract_folder = self.download_folder / rar_path.stem
                        counter = 1
                        original_extract_folder = extract_folder
                        while extract_folder.exists():
                            extract_folder = Path(f"{original_extract_folder}_{counter}")
                            counter += 1
                        rar_ref.extractall(str(extract_folder))
            except rarfile.NeedFirstVolume:
                self.log(f"Error: {rar_path.name} is a multi-part RAR archive. Please provide all parts.")
                return
            except rarfile.Error as e:
                self.log(f"Error: Could not extract {rar_path.name}: {e}")
                return
            except Exception as e:
                self.log(f"Error: Could not open/extract {rar_path.name}: {e}")
                return
            if extract_to_downloads:
                self.log(f"Successfully extracted to monitored folder: {extract_folder}")
                search_folder = extract_folder
            else:
                self.log(f"Successfully extracted to: {extract_folder}")
                search_folder = extract_folder
            # Copy whole folder if enabled
            if self.copy_whole_folder:
                self.log("Copying entire extracted folder to destination (option enabled)...")
                self._copy_entire_folder(search_folder)
                # Always try to delete after copying (handled in _copy_entire_folder)
            else:
                self.copy_selected_files(search_folder, stop_event=stop_event)
            self.processed_files.add(rar_path)
            if self.delete_after_zip:
                try:
                    if rar_path.exists():
                        rar_path.unlink()
                        self.log(f"Deleted RAR file: {rar_path}")
                except Exception as e:
                    self.log(f"Failed to delete RAR file: {rar_path} ({e})")
        except rarfile.BadRarFile:
            self.log(f"Error: {rar_path.name} is not a valid RAR file or is corrupted")
        except PermissionError:
            self.log(f"Error: Permission denied accessing {rar_path.name}")
        except Exception as e:
            self.log(f"Error extracting {rar_path.name}: {str(e)}")

    def copy_selected_files(self, folder, stop_event=None):
        deleted = False
        copied_any = False
        # Each option works independently, both can copy files if both are enabled
        if self.logic_enabled and self.logic_input:
            copied_any = self._copy_files_with_priority_logic(folder, stop_event=stop_event) or copied_any
        if self.copy_enabled:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if stop_event and stop_event.is_set():
                        self.log("Copying stopped by user.")
                        return deleted
                    ext = Path(file).suffix.lower().lstrip(".")
                    if self.collect_exts is None or ext in self.collect_exts:
                        src_file = Path(root) / file
                        dest_file = self.target_folder / file
                        counter = 1
                        base_name = dest_file.stem
                        ext_name = dest_file.suffix
                        while dest_file.exists():
                            dest_file = self.target_folder / f"{base_name}_{counter}{ext_name}"
                            counter += 1
                        try:
                            copy2(src_file, dest_file)
                            self.log(f"Copied: {src_file} -> {dest_file}")
                            copied_any = True
                        except Exception as e:
                            self.log(f"Failed to copy {src_file}: {e}")
        if (self.copy_enabled or (self.logic_enabled and self.logic_input)) and self.delete_after_extracted and Path(folder).exists() and copied_any:
            try:
                import shutil
                shutil.rmtree(folder)
                self.log(f"Deleted extracted folder after copying: {folder}")
                deleted = True
            except Exception as e:
                self.log(f"Failed to delete extracted folder after copying: {folder} ({e})")
        if not self.copy_enabled and not (self.logic_enabled and self.logic_input):
            self.log("Copying skipped (option not selected).")
        return deleted

    def _copy_files_with_priority_logic(self, folder, stop_event=None):
        logic_input = self.logic_input or ""
        priorities = []
        for part in logic_input.split(";"):
            part = part.strip()
            if not part:
                continue
            # Remove priority number if present
            if "-" in part:
                _, exts = part.split("-", 1)
            else:
                exts = part
            ext_list = [e.strip().lstrip(".").lower() for e in exts.split(",") if e.strip()]
            if ext_list:
                priorities.append(ext_list)
        if not priorities:
            self.log("No valid logic found in input.")
            return False

        self.log(f"Priority logic parsed: {priorities}")

        # Gather all files in folder (recursively)
        all_files = []
        for root, dirs, files in os.walk(folder):
            for file in files:
                all_files.append(Path(root) / file)

        # Try each priority group in order
        for idx, ext_group in enumerate(priorities):
            matched_files = []
            for f in all_files:
                if stop_event and stop_event.is_set():
                    self.log("Copying stopped by user.")
                    return False
                ext = f.suffix.lower().lstrip(".")
                if ext in ext_group:
                    matched_files.append(f)
            if matched_files:
                self.log(f"Priority {idx+1}: Found files with extensions {ext_group}:")
                for src_file in matched_files:
                    if stop_event and stop_event.is_set():
                        self.log("Copying stopped by user.")
                        return False
                    dest_file = self.target_folder / src_file.name
                    counter = 1
                    base_name = dest_file.stem
                    ext_name = dest_file.suffix
                    while dest_file.exists():
                        dest_file = self.target_folder / f"{base_name}_{counter}{ext_name}"
                        counter += 1
                    try:
                        copy2(src_file, dest_file)
                        self.log(f"Copied: {src_file} -> {dest_file}")
                    except Exception as e:
                        self.log(f"Failed to copy {src_file}: {e}")
                self.log(f"Stopped at priority {idx+1}, no lower priorities will be checked.")
                return True
            else:
                self.log(f"Priority {idx+1}: No files found for extensions {ext_group}.")
        self.log("No files matched any priority group. Nothing copied.")
        return False

    def _copy_entire_folder(self, src_folder):
        import shutil
        dest = self.target_folder / Path(src_folder).name
        counter = 1
        orig_dest = dest
        while dest.exists():
            dest = self.target_folder / f"{orig_dest.stem}_{counter}{orig_dest.suffix}"
            counter += 1
        try:
            shutil.copytree(src_folder, dest)
            self.log(f"Copied entire folder: {src_folder} -> {dest}")
            # Always delete extracted folder after copying if option is enabled
            if self.delete_after_extracted and Path(src_folder).exists():
                try:
                    shutil.rmtree(src_folder)
                    self.log(f"Deleted extracted folder after copying: {src_folder}")
                except Exception as e:
                    self.log(f"Failed to delete extracted folder after copying: {src_folder} ({e})")
        except Exception as e:
            self.log(f"Failed to copy entire folder: {src_folder} -> {dest}: {e}")

class UnzipperGUI:
    def __init__(self, root):
        self.root = root
        # --- Set window icon for taskbar (works for .ico only) ---
        import sys
        import os
        def get_icon_path():
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller bundle
                return os.path.join(sys._MEIPASS, 'icon.ico')
            else:
                return str(get_base_dir() / 'icon.ico')
        icon_path = get_icon_path()
        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            pass  # Fallback: no icon

        # --- REDESIGN: Clean, modern, non-transparent UI ---
        # Remove transparency and set a solid background
        self.root.configure(bg="#f4f6fb")
        self.root.geometry("920x620")
        self.root.minsize(920, 620)
        self.root.title("Unzipper")
        # Remove any transparency attributes
        try:
            self.root.wm_attributes("-transparentcolor", "")
        except Exception:
            pass

        # Main content frame
        main_frame = tk.Frame(root, bg="#ffffff", bd=0, highlightbackground="#ffffff", highlightthickness=0)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 16))

        # Section style helper for consistent border and padding
        # Remove border/highlight from section frames, increase vertical padding
        def section_frame(parent):
            f = tk.Frame(parent, bg="#ffffff", highlightthickness=0, bd=0)
            f.pack(fill=None, anchor="w", padx=8, pady=10, ipadx=0, ipady=0)
            return f

        # Folder selection
        folder_frame = section_frame(main_frame)
        tk.Label(folder_frame, text="Monitor folder:", bg="#ffffff", fg="#333", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        self.monitor_var = tk.StringVar()
        self.monitor_entry = tk.Entry(folder_frame, textvariable=self.monitor_var, width=73, font=("Segoe UI", 11), bg="#f7fafd", relief="flat", highlightthickness=1, highlightbackground="#bdbdbd")
        self.monitor_entry.pack(side=tk.LEFT, padx=(0, 8))
        self.monitor_select_btn = tk.Button(folder_frame, text="Select", command=self.select_monitor_folder, height=1)
        self.monitor_select_btn.pack(side=tk.LEFT)
        # Set a fixed min height for the select button to match entry height
        self.monitor_select_btn.configure(height=1)

        folder2_frame = section_frame(main_frame)
        tk.Label(folder2_frame, text="Destination folder:", bg="#ffffff", fg="#333", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        self.dest_var = tk.StringVar()
        self.dest_entry = tk.Entry(folder2_frame, textvariable=self.dest_var, width=70, font=("Segoe UI", 11), bg="#f7fafd", relief="flat", highlightthickness=1, highlightbackground="#bdbdbd")
        self.dest_entry.pack(side=tk.LEFT, padx=(0, 8))
        self.dest_select_btn = tk.Button(folder2_frame, text="Select", command=self.select_dest_folder, height=1)
        self.dest_select_btn.pack(side=tk.LEFT)
        # Set a fixed min height for the select button to match entry height
        self.dest_select_btn.configure(height=1)

        # File extension and logic options
        ext_logic_frame = section_frame(main_frame)
        # Use a grid layout for better alignment and control
        ext_logic_frame.grid_columnconfigure(1, weight=1)
        ext_logic_frame.grid_columnconfigure(2, weight=1)
        ext_logic_frame.grid_columnconfigure(3, weight=1)
        ext_logic_frame.grid_columnconfigure(4, weight=1)
        ext_logic_frame.grid_columnconfigure(5, weight=1)
        ext_logic_frame.grid_columnconfigure(6, weight=1)
        ext_logic_frame.grid_columnconfigure(7, weight=1)
        ext_logic_frame.grid_columnconfigure(8, weight=1)

        # --- First row: Copy files ---
        self.copy_enabled_var = tk.BooleanVar(value=True)
        self.copy_chk = tk.Checkbutton(ext_logic_frame, text="Copy files", variable=self.copy_enabled_var, bg="#ffffff", font=("Segoe UI", 11), activebackground="#e3f2fd", selectcolor="#e3f2fd", command=self.on_copy_enabled_changed)
        self.copy_chk.grid(row=0, column=0, padx=(0, 8), pady=(0, 2), sticky="w")
        tk.Label(ext_logic_frame, text="Extensions:", bg="#ffffff", fg="#333", font=("Segoe UI", 11)).grid(row=0, column=1, sticky="e")
        self.ext_var = tk.StringVar()
        self.ext_entry = tk.Entry(ext_logic_frame, textvariable=self.ext_var, width=70, font=("Segoe UI", 11), bg="#f7fafd", relief="flat", highlightthickness=1, highlightbackground="#bdbdbd")
        self.ext_entry.grid(row=0, column=2, columnspan=1, padx=(0, 12), sticky="ew")
        # Add tooltip for copy files input
        def show_ext_tip(event):
            x = event.x_root + 10
            y = event.y_root + 10
            self.ext_tip = tk.Toplevel(self.ext_entry)
            self.ext_tip.wm_overrideredirect(True)
            self.ext_tip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                self.ext_tip,
                text="Example: jpg, png, ai",
                background="#ffffe0",
                relief='solid',
                borderwidth=1,
                font=("Segoe UI", 9, "normal"),
                justify="left"
            )
            label.pack(ipadx=1)
        def hide_ext_tip(event):
            if hasattr(self, "ext_tip"):
                self.ext_tip.destroy()
                del self.ext_tip
        self.ext_entry.bind("<Enter>", show_ext_tip)
        self.ext_entry.bind("<Leave>", hide_ext_tip)

        # --- Second row: Copy files with logic ---
        self.copy_logic_enabled_var = tk.BooleanVar(value=False)
        self.copy_logic_chk = tk.Checkbutton(ext_logic_frame, text="Copy files with logic", variable=self.copy_logic_enabled_var, bg="#ffffff", font=("Segoe UI", 11), activebackground="#e3f2fd", selectcolor="#e3f2fd", command=self.on_copy_logic_enabled_changed)
        self.copy_logic_chk.grid(row=1, column=0, padx=(0, 8), pady=(2, 0), sticky="w")
        tk.Label(ext_logic_frame, text="Logic:", bg="#ffffff", fg="#333", font=("Segoe UI", 11)).grid(row=1, column=1, sticky="e")
        self.copy_logic_var = tk.StringVar()
        self.copy_logic_entry = tk.Entry(ext_logic_frame, textvariable=self.copy_logic_var, width=60, font=("Segoe UI", 11), bg="#f7fafd", relief="flat", highlightthickness=1, highlightbackground="#bdbdbd")
        self.copy_logic_entry.grid(row=1, column=2, columnspan=1, sticky="ew")
        # Add tooltip for logic input
        def show_tip(event):
            x = event.x_root + 10
            y = event.y_root + 10
            self.logic_tip = tk.Toplevel(self.copy_logic_entry)
            self.logic_tip.wm_overrideredirect(True)
            self.logic_tip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                self.logic_tip,
                text="Priority based logic: If higher priority found only it is copied.\n\nExample: ai; png, esp; jpg\n\nLeft most entry is the higest priority.\nSemicolon separates different priority level\nComma separate extensions with same priority.",
                background="#ffffe0",
                relief='solid',
                borderwidth=1,
                font=("Segoe UI", 9, "normal"),
                justify="left"
            )
            label.pack(ipadx=1)
        def hide_tip(event):
            if hasattr(self, "logic_tip"):
                self.logic_tip.destroy()
                del self.logic_tip
        self.copy_logic_entry.bind("<Enter>", show_tip)
        self.copy_logic_entry.bind("<Leave>", hide_tip)

        # --- Third row: Copy whole extracted folder ---
        self.copy_whole_folder_var = tk.BooleanVar(value=False)
        self.copy_whole_folder_chk = tk.Checkbutton(
            ext_logic_frame,
            text="Copy whole extracted folder to destination",
            variable=self.copy_whole_folder_var,
            bg="#ffffff",
            font=("Segoe UI", 11),
            activebackground="#e3f2fd",
            selectcolor="#e3f2fd",
            anchor="w",
            command=self.on_copy_whole_folder_changed
        )
        self.copy_whole_folder_chk.grid(row=2, column=0, padx=(0, 8), pady=(2, 0), sticky="w", columnspan=3)

        # Delete and startup options
        options_frame = section_frame(main_frame)
        self.delete_zip_var = tk.BooleanVar()
        self.delete_extracted_var = tk.BooleanVar()
        self.delete_zip_chk = tk.Checkbutton(options_frame, text="Delete ZIP/RAR after extracting", variable=self.delete_zip_var, bg="#ffffff", font=("Segoe UI", 11), activebackground="#e3f2fd", selectcolor="#e3f2fd", command=self.on_delete_zip_changed)
        self.delete_zip_chk.pack(side=tk.LEFT, padx=(0, 12))
        self.delete_extracted_chk = tk.Checkbutton(options_frame, text="Delete extracted folder after copying", variable=self.delete_extracted_var, bg="#ffffff", font=("Segoe UI", 11), activebackground="#e3f2fd", selectcolor="#e3f2fd", command=self.on_delete_extracted_changed)
        self.delete_extracted_chk.pack(side=tk.LEFT, padx=(0, 12))
        self.startup_var = tk.BooleanVar(value=is_startup_enabled())
        self.startup_chk = tk.Checkbutton(options_frame, text="Run at startup", variable=self.startup_var, bg="#ffffff", font=("Segoe UI", 11), activebackground="#e3f2fd", selectcolor="#e3f2fd", command=self.toggle_startup)
        self.startup_chk.pack(side=tk.LEFT, padx=(0, 12))
        self.tray_btn = tk.Button(options_frame, text="Minimize to Tray", command=self.hide_window_to_tray)
        self.tray_btn.pack(side=tk.LEFT, padx=(0, 8))

        # Action buttons
        btn_frame = section_frame(main_frame)
        self.start_btn = tk.Button(btn_frame, text="Start Monitoring", command=self.start_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stop_btn = tk.Button(btn_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        save_btn = tk.Button(btn_frame, text="Save Config", command=self.save_config)
        save_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._extract_all_stop_event = threading.Event()
        self.extract_all_btn = tk.Button(btn_frame, text="Extract All Existing", command=self.extract_all_archives)
        self.extract_all_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stop_extract_all_btn = tk.Button(btn_frame, text="Stop Extracting", command=self.stop_extract_all, state=tk.DISABLED)
        self.stop_extract_all_btn.pack(side=tk.LEFT, padx=(0, 8))

        # Log area
        self.log_area = scrolledtext.ScrolledText(main_frame, state='disabled', height=12, font=("Consolas", 11), bg="#f7fafd", fg="#222", relief="flat", highlightthickness=1, highlightbackground="#bdbdbd", bd=0)
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=(8, 8), padx=8)  # Add left/right and bottom padding

        # Define log method before any code that may call it
        def log(msg):
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, msg + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        self.log = log

        # Only bind close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Load config if exists
        self.load_config()

        # Set initial state for both entries after both are created
        self.on_copy_enabled_changed()
        self.on_copy_logic_enabled_changed()

        # --- Modern UI Styling ---
        # Remove duplicate title bar at the bottom (if present)
        # (No code for a second title bar or label should be here)

        # Modernize all frames
        def style_frame(frm):
            frm.configure(bg="#ffffff", highlightbackground="#e0e0e0", highlightthickness=1, bd=0)
        style_frame(main_frame)

        # Style all entries
        entry_style = {"bg": "#f7fafd", "relief": "flat", "highlightthickness": 1, "highlightbackground": "#bdbdbd", "bd": 0}
        self.monitor_entry.config(**entry_style)
        self.dest_entry.config(**entry_style)
        self.ext_entry.config(**entry_style)
        self.copy_logic_entry.config(**entry_style)

        # Style all buttons
        def style_button(btn, small=False):
            # Only override height if not a select button
            if btn in [self.monitor_select_btn, self.dest_select_btn]:
                btn.config(relief="solid", bd=1, padx=10, pady=2, font=("Segoe UI", 10 if small else 11, "bold"), highlightbackground="#bdbdbd", highlightthickness=2, height=1)
            else:
                btn.config(relief="solid", bd=1, padx=10, pady=2, font=("Segoe UI", 10 if small else 11, "bold"), highlightbackground="#bdbdbd", highlightthickness=2)
        for btn in [self.start_btn, self.stop_btn, self.extract_all_btn, self.stop_extract_all_btn, self.tray_btn, self.monitor_select_btn, self.dest_select_btn]:
            style_button(btn, small=True)
        for child in btn_frame.winfo_children():
            if isinstance(child, tk.Button) and child not in [self.start_btn, self.stop_btn, self.extract_all_btn, self.stop_extract_all_btn, self.tray_btn, self.monitor_select_btn, self.dest_select_btn]:
                style_button(child, small=True)
        # Remove hover effect coloring
        def on_enter(e): pass
        def on_leave(e): pass
        for btn in [self.start_btn, self.stop_btn, self.extract_all_btn, self.stop_extract_all_btn, self.tray_btn]:
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
        for child in btn_frame.winfo_children():
            if isinstance(child, tk.Button) and child not in [self.start_btn, self.stop_btn, self.extract_all_btn, self.stop_extract_all_btn, self.tray_btn]:
                child.bind("<Enter>", on_enter)
                child.bind("<Leave>", on_leave)

        # Remove outer border and padding for seamless look
        main_frame.config(bd=0, highlightthickness=0, highlightbackground="#ffffff")
        main_frame.pack_configure(padx=0, pady=0)

        # Style checkboxes
        for chk in [self.copy_chk, self.copy_logic_chk, self.copy_whole_folder_chk, self.delete_zip_chk, self.delete_extracted_chk, self.startup_chk]:
            chk.config(bg="#ffffff", activebackground="#e3f2fd", selectcolor="#e3f2fd", font=("Segoe UI", 11))

        # Style labels
        for f in [main_frame, folder_frame, folder2_frame, ext_logic_frame, options_frame]:
            for child in f.winfo_children():
                if isinstance(child, tk.Label):
                    child.config(bg="#ffffff", font=("Segoe UI", 11, "bold"), fg="#333")

        # Style log area
        self.log_area.config(bg="#f7fafd", fg="#222", font=("Consolas", 11), relief="flat", highlightthickness=1, highlightbackground="#bdbdbd", bd=0)

        # Ensure tray_icon is always defined
        self.tray_icon = None

        # Start monitoring automatically after UI setup and config load
        self.start_monitoring()

    def restart_monitoring(self):
        # Stop and restart monitoring if currently running
        if hasattr(self, 'monitoring') and self.monitoring:
            self.stop_monitoring()
            self.start_monitoring()

    def select_monitor_folder(self):
        folder = filedialog.askdirectory(title="Select folder to monitor")
        if folder:
            self.monitor_var.set(folder)
            self.restart_monitoring()

    def select_dest_folder(self):
        folder = filedialog.askdirectory(title="Select destination folder")
        if folder:
            self.dest_var.set(folder)
            self.restart_monitoring()

    def save_config(self):
        monitor_folder = self.monitor_var.get()
        dest_folder = self.dest_var.get()
        delete_after_zip = self.delete_zip_var.get()
        delete_after_extracted = self.delete_extracted_var.get()
        file_exts = self.ext_var.get()
        logic_input = self.copy_logic_var.get()
        copy_enabled = self.copy_enabled_var.get()
        logic_enabled = self.copy_logic_enabled_var.get()
        copy_whole_folder = self.copy_whole_folder_var.get()
        if not monitor_folder or not dest_folder:
            messagebox.showerror("Error", "Both folders must be selected.")
            return
        write_config(
            monitor_folder, dest_folder, delete_after_zip, delete_after_extracted,
            file_exts, logic_input,
            copy_enabled=copy_enabled,
            logic_enabled=logic_enabled,
            copy_whole_folder=copy_whole_folder
        )
        self.log("Configuration saved.")

    def load_config(self):
        config = read_config()
        if "monitor_folder" in config:
            self.monitor_var.set(config["monitor_folder"])
        if "dest_folder" in config:
            self.dest_var.set(config["dest_folder"])
        if "delete_after_zip" in config:
            self.delete_zip_var.set(config["delete_after_zip"].lower() == "true")
        if "delete_after_extracted" in config:
            self.delete_extracted_var.set(config["delete_after_extracted"].lower() == "true")
        if "file_exts" in config:
            self.ext_var.set(config["file_exts"])
        if "logic_input" in config:
            self.copy_logic_var.set(config["logic_input"])
        if "copy_enabled" in config:
            self.copy_enabled_var.set(config["copy_enabled"].lower() == "true")
        if "logic_enabled" in config:
            self.copy_logic_enabled_var.set(config["logic_enabled"].lower() == "true")
        if "copy_whole_folder" in config:
            self.copy_whole_folder_var.set(config["copy_whole_folder"].lower() == "true")

    def start_monitoring(self):
        monitor_folder = self.monitor_var.get()
        dest_folder = self.dest_var.get()
        delete_after_zip = self.delete_zip_var.get()
        delete_after_extracted = self.delete_extracted_var.get()
        file_exts = self.ext_var.get()
        copy_enabled = self.copy_enabled_var.get()
        logic_enabled = self.copy_logic_enabled_var.get()
        logic_input = self.copy_logic_var.get()
        copy_whole_folder = self.copy_whole_folder_var.get()
        if not monitor_folder or not dest_folder:
            messagebox.showerror("Error", "Both folders must be selected.")
            return
        if not Path(monitor_folder).exists() or not Path(dest_folder).exists():
            messagebox.showerror("Error", "Selected folders do not exist.")
            return
        self.save_config()
        self.handler = ZipExtractorHandler(
            monitor_folder, dest_folder,
            delete_after_zip=delete_after_zip,
            delete_after_extracted=delete_after_extracted,
            file_exts=file_exts,
            gui_callback=self.log,
            copy_enabled=copy_enabled,
            logic_input=logic_input,
            logic_enabled=logic_enabled,
            copy_whole_folder=copy_whole_folder
        )
        self.observer = Observer()
        self.observer.schedule(self.handler, str(monitor_folder), recursive=False)
        self.monitoring = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.log(f"Monitoring folder: {monitor_folder}")
        self.log("Watching for new ZIP files... (Press Stop Monitoring to stop)")
        threading.Thread(target=self._run_observer, daemon=True).start()

    def _run_observer(self):
        self.observer.start()
        try:
            while self.monitoring:
                time.sleep(1)
        except Exception as e:
            self.log(f"Error: {e}")
        finally:
            self.observer.stop()
            self.observer.join()
            self.log("ZIP file monitor stopped.")

    def stop_monitoring(self):
        self.monitoring = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.log("Stopping ZIP file monitor...")

    def toggle_startup(self):
        if self.startup_var.get():
            ok = create_startup_shortcut()
            if ok:
                self.log("Startup shortcut created.")
            else:
                self.log("Failed to create startup shortcut (Windows only, requires pywin32 and winshell).")
        else:
            if remove_startup_shortcut():
                self.log("Startup shortcut removed.")
            else:
                self.log("No startup shortcut to remove.")

    def hide_window_to_tray(self):
        if not pystray or self.tray_icon:
            return
        self.root.withdraw()
        self.create_tray_icon()

    def show_window_from_tray(self, icon=None, item=None):
        self.root.after(0, self._show_window)

    def _show_window(self):
        self.root.deiconify()
        self.root.after(0, self.root.lift)
        self.root.after(0, lambda: self.root.focus_force())
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

    def create_tray_icon(self):
        if not pystray:
            return
        import sys
        import os
        def get_icon_path():
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, 'icon.ico')
            else:
                return str(get_base_dir() / 'icon.ico')
        icon_path = get_icon_path()
        image = None
        # Prefer .ico for tray icon if available
        if os.path.exists(icon_path):
            try:
                image = Image.open(icon_path)
            except Exception:
                image = None
        if image is None:
            # fallback: blue square
            image = Image.new('RGB', (64, 64), color=(0, 120, 215))
            d = ImageDraw.Draw(image)
            d.rectangle([16, 16, 48, 48], fill=(255, 255, 255))
        menu = pystray.Menu(
            pystray.MenuItem('Restore', self.show_window_from_tray),
            pystray.MenuItem('Exit', self.exit_from_tray)
        )
        self.tray_icon = pystray.Icon("Unzipper", image, "Unzipper", menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()

    def exit_from_tray(self, icon=None, item=None):
        self.root.after(0, self._exit_app)

    def _exit_app(self):
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.root.destroy()

    def on_close(self):
        # Only minimize to tray if already hidden, otherwise exit
        if self.tray_icon:
            self.root.withdraw()
        else:
            self._exit_app()

    def on_copy_enabled_changed(self):
        if self.copy_enabled_var.get():
            self.ext_entry.config(state='normal')
        else:
            self.ext_entry.config(state='disabled')
        self.restart_monitoring()

    def on_copy_logic_enabled_changed(self):
        if self.copy_logic_enabled_var.get():
            self.copy_logic_entry.config(state='normal')
        else:
            self.copy_logic_entry.config(state='disabled')
        self.restart_monitoring()

    def on_copy_whole_folder_changed(self):
        self.restart_monitoring()

    def on_delete_zip_changed(self):
        self.restart_monitoring()

    def on_delete_extracted_changed(self):
        self.restart_monitoring()

    def on_copy_logic_apply(self):
        # Placeholder for logic to be implemented later
        messagebox.showinfo("Info", "Copy files with logic will be implemented later.")

    def extract_all_archives(self):
        self._extract_all_stop_event.clear()
        self.extract_all_btn.config(state=tk.DISABLED)
        self.stop_extract_all_btn.config(state=tk.NORMAL)
        monitor_folder = self.monitor_var.get()
        dest_folder = self.dest_var.get()
        delete_after_zip = self.delete_zip_var.get()
        delete_after_extracted = self.delete_extracted_var.get()
        file_exts = self.ext_var.get()
        copy_enabled = self.copy_enabled_var.get()
        logic_enabled = self.copy_logic_enabled_var.get()
        logic_input = self.copy_logic_var.get()
        copy_whole_folder = self.copy_whole_folder_var.get()
        if not monitor_folder or not dest_folder:
            self.log("Both folders must be selected.")
            self.extract_all_btn.config(state=tk.NORMAL)
            self.stop_extract_all_btn.config(state=tk.DISABLED)
            return
        if not Path(monitor_folder).exists() or not Path(dest_folder).exists():
            self.log("Selected folders do not exist.")
            self.extract_all_btn.config(state=tk.NORMAL)
            self.stop_extract_all_btn.config(state=tk.DISABLED)
            return

        def do_extract():
            try:
                handler = ZipExtractorHandler(
                    monitor_folder, dest_folder,
                    delete_after_zip=delete_after_zip,
                    delete_after_extracted=delete_after_extracted,
                    file_exts=file_exts,
                    gui_callback=self.log,
                    copy_enabled=copy_enabled,
                    logic_input=logic_input,
                    logic_enabled=logic_enabled,
                    copy_whole_folder=copy_whole_folder
                )
                archive_exts = handler.archive_exts
                monitor_path = Path(monitor_folder)
                archive_files = [f for f in monitor_path.iterdir() if f.is_file() and f.suffix.lower() in archive_exts]
                if not archive_files:
                    self.log("No ZIP or RAR files found to extract.")
                    self.extract_all_btn.config(state=tk.NORMAL)
                    self.stop_extract_all_btn.config(state=tk.DISABLED)
                    return

                self.log(f"Extracting {len(archive_files)} archive(s)...")
                for file_path in archive_files:
                    if self._extract_all_stop_event.is_set():
                        self.log("Extraction stopped by user.")
                        break
                    ext = file_path.suffix.lower()
                    try:
                        if ext == '.zip':
                            if self._extract_all_stop_event.is_set():
                                self.log(f"Stopped before extracting {file_path.name}.")
                                break
                            handler.extract_zip(file_path, stop_event=self._extract_all_stop_event)
                        elif ext == '.rar' and rarfile:
                            if self._extract_all_stop_event.is_set():
                                self.log(f"Stopped before extracting {file_path.name}.")
                                break
                            handler.extract_rar(file_path, stop_event=self._extract_all_stop_event)
                    except Exception as e:
                        self.log(f"Error extracting {file_path}: {e}")
                    if self._extract_all_stop_event.is_set():
                        self.log("Extraction stopped by user.")
                        break
                else:
                    self.log("Extraction of all archives complete.")
            except Exception as e:
                self.log(f"Unexpected error during extraction: {e}")
            finally:
                self.extract_all_btn.config(state=tk.NORMAL)
                self.stop_extract_all_btn.config(state=tk.DISABLED)

        threading.Thread(target=do_extract, daemon=True).start()

    def stop_extract_all(self):
        self._extract_all_stop_event.set()
        self.log("Stopping extraction immediately...")

def main():
    root = tk.Tk()
    app = UnzipperGUI(root)
    app.start_monitoring()
    root.mainloop()

if __name__ == "__main__":
    main()
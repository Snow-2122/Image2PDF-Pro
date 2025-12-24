import os
import shutil
import re
import sys
import logging
import signal
import gc
import zipfile
import io
import configparser
import threading
import concurrent.futures
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import ctypes
import webbrowser
from PIL import Image, ImageTk
from tqdm import tqdm

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import ToolTip
from plyer import notification

try:
    import winsound
except ImportError:
    winsound = None

try:
    import rarfile
except ImportError:
    rarfile = None

# --- High DPI Fix ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# --- Path Locking (Critical for .exe and Drag-and-Drop) ---
if getattr(sys, 'frozen', False):
    # If run as .exe
    APP_DIR = os.path.dirname(sys.executable)
else:
    # If run as script
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, 'config.ini')
LOG_FILE = os.path.join(APP_DIR, 'errors.log')

# --- Configuration ---
def load_config():
    config = configparser.ConfigParser()
    defaults = {
        'PDF_Quality': 100,
        'Image_Extensions': {'.jpg', '.jpeg', '.png', '.webp', '.bmp'},
        'Parallel_Processing': True,
        'Thread_Count': 4,
        'Enable_GUI': True,
        'Enable_RAR': True,
        'Theme': 'solar',
        'Landscape_Mode': 'none',
        'Output_Path': '',
        'Delete_Source': False
    }
    
    if not os.path.exists(CONFIG_FILE):
        return defaults
    
    config.read(CONFIG_FILE)
    if 'General' not in config: return defaults
    general = config['General']
    
    ext_str = general.get('Image_Extensions', '.jpg, .jpeg, .png, .webp, .bmp')
    extensions = {e.strip() for e in ext_str.split(',')}
    
    return {
        'PDF_Quality': general.getint('PDF_Quality', 100),
        'Image_Extensions': extensions,
        'Parallel_Processing': general.getboolean('Parallel_Processing', True),
        'Thread_Count': general.getint('Thread_Count', 4),
        'Enable_GUI': general.getboolean('Enable_GUI', True),
        'Enable_RAR': general.getboolean('Enable_RAR', True),
        'Theme': general.get('Theme', 'solar'),
        'Landscape_Mode': general.get('Landscape_Mode', 'none'),
        'Output_Path': general.get('Output_Path', ''),
        'Delete_Source': general.getboolean('Delete_Source', False)
    }

def save_config(settings):
    config = configparser.ConfigParser()
    config['General'] = {
        'PDF_Quality': str(settings['PDF_Quality']),
        'Image_Extensions': ', '.join(settings['Image_Extensions']),
        'Parallel_Processing': str(settings['Parallel_Processing']),
        'Thread_Count': str(settings['Thread_Count']),
        'Enable_GUI': str(settings['Enable_GUI']),
        'Enable_RAR': str(settings['Enable_RAR']),
        'Theme': settings['Theme'],
        'Landscape_Mode': settings['Landscape_Mode'],
        'Output_Path': settings['Output_Path'],
        'Delete_Source': str(settings['Delete_Source'])
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

SETTINGS = load_config()

# Setup logging
logging.basicConfig(
    filename=LOG_FILE, 
    level=logging.ERROR, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

ABORT_REQUESTED = False

def signal_handler(signum, frame):
    global ABORT_REQUESTED
    print("\n\nAborting... Please wait for cleanup.")
    ABORT_REQUESTED = True

signal.signal(signal.SIGINT, signal_handler)

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split('([0-9]+)', s)]

def optimize_image(img, mode='none'):
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    width, height = img.size
    
    # Check if Landscape
    if width > height:
        if mode == 'none':
            # Keep original aspect ratio (mixed orientation)
            return [img.copy()]
            
        elif mode == 'rotate':
            return [img.rotate(270, expand=True)]
        
        elif mode == 'split':
            half_width = width // 2
            right_box = (half_width, 0, width, height)
            right_part = img.crop(right_box)
            left_box = (0, 0, half_width, height)
            left_part = img.crop(left_box)
            return [right_part, left_part]

        else: # letterbox
            new_height = int(width * 1.3)
            new_bg = Image.new('RGB', (width, new_height), (255, 255, 255))
            y_offset = (new_height - height) // 2
            new_bg.paste(img, (0, y_offset))
            return [new_bg]
    else:
        return [img.copy()]

def process_images_to_pdf(image_list, pdf_path, title):
    if not image_list: return False
    try:
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        base_image = image_list[0]
        other_images = image_list[1:]
        # Enforce Lossless Quality
        base_image.save(
            pdf_path, "PDF", resolution=100.0, save_all=True, append_images=other_images,
            quality=100, subsampling=0, title=title, producer="Image2PDF Script"
        )
        return True
    except Exception as e:
        logging.error(f"Error saving PDF {pdf_path}: {e}")
        return False
    finally:
        for img in image_list:
            if hasattr(img, 'close'): img.close()

def get_pdf_path(original_path, source_base_path=None):
    filename = f"{os.path.splitext(os.path.basename(original_path))[0]}.pdf"
    custom_out = SETTINGS.get('Output_Path', '').strip()
    if custom_out and os.path.exists(custom_out):
        return os.path.join(custom_out, filename)
    return os.path.join(os.path.dirname(original_path), filename)

def process_folder(folder_path, pdf_path, progress_callback=None):
    if ABORT_REQUESTED: return False
    try:
        files = os.listdir(folder_path)
        image_files = [f for f in files if os.path.splitext(f.lower())[1] in SETTINGS['Image_Extensions']]
        if not image_files: return False

        image_files.sort(key=natural_sort_key)
        images = []
        
        for i, filename in enumerate(image_files):
            if ABORT_REQUESTED: return False
            try:
                with Image.open(os.path.join(folder_path, filename)) as img:
                    processed_imgs = optimize_image(img, mode=SETTINGS['Landscape_Mode'])
                    images.extend(processed_imgs)
            except Exception as e:
                logging.error(f"Error {filename}: {e}")
            if progress_callback: progress_callback(i + 1, len(image_files))

        if not images: return False
        success = process_images_to_pdf(images, pdf_path, os.path.basename(folder_path))
        del images
        return success
    except Exception as e:
        logging.error(f"Folder error {folder_path}: {e}")
        return False

def process_archive(archive_path, pdf_path, progress_callback=None):
    if ABORT_REQUESTED: return False
    ext = os.path.splitext(archive_path.lower())[1]
    try:
        files_data = []
        if ext in ['.zip', '.cbz']:
            with zipfile.ZipFile(archive_path, 'r') as zf:
                names = zf.namelist()
                valid_names = [n for n in names if os.path.splitext(n.lower())[1] in SETTINGS['Image_Extensions']]
                valid_names.sort(key=natural_sort_key)
                for i, name in enumerate(valid_names):
                    if ABORT_REQUESTED: return False
                    files_data.append(zf.read(name))
                    if progress_callback: progress_callback(i, len(valid_names))
        elif ext in ['.rar', '.cbr'] and SETTINGS['Enable_RAR']:
            if not rarfile: raise ImportError("rarfile module missing")
            with rarfile.RarFile(archive_path, 'r') as rf:
                names = rf.namelist()
                valid_names = [n for n in names if os.path.splitext(n.lower())[1] in SETTINGS['Image_Extensions']]
                valid_names.sort(key=natural_sort_key)
                for i, name in enumerate(valid_names):
                    if ABORT_REQUESTED: return False
                    files_data.append(rf.read(name))
                    if progress_callback: progress_callback(i, len(valid_names))

        images = []
        for i, data in enumerate(files_data):
            if ABORT_REQUESTED: return False
            try:
                with Image.open(io.BytesIO(data)) as img:
                    processed_imgs = optimize_image(img, mode=SETTINGS['Landscape_Mode'])
                    images.extend(processed_imgs)
            except Exception as e:
                logging.error(f"Image error in archive: {e}")
            if progress_callback: progress_callback(i, len(files_data))

        if not images: return False
        success = process_images_to_pdf(images, pdf_path, os.path.basename(archive_path))
        del images
        return success
    except Exception as e:
        logging.error(f"Archive error {archive_path}: {e}")
        return False

def worker_task(item):
    if ABORT_REQUESTED: return False
    success = False
    if item['type'] == 'folder':
        success = process_folder(item['path'], item['pdf_path'])
        if success and SETTINGS['Delete_Source']: shutil.rmtree(item['path'], ignore_errors=True)
    elif item['type'] == 'archive':
        success = process_archive(item['path'], item['pdf_path'])
        if success and SETTINGS['Delete_Source']: os.remove(item['path'])
    gc.collect()
    return success

# --- UnRAR Check ---
def check_unrar_status():
    """Returns 'OK', 'COPY', or 'DOWNLOAD'."""
    # 1. Check Local (Best)
    local_unrar = os.path.join(APP_DIR, "UnRAR.exe")
    if os.path.exists(local_unrar):
        if rarfile: rarfile.UNRAR_TOOL = local_unrar
        return "OK"
    
    # 2. Check System PATH
    if shutil.which("UnRAR") or shutil.which("UnRAR.exe"):
        return "OK"

    # 3. Check Program Files (Common WinRAR path)
    winrar_path = r"C:\Program Files\WinRAR\UnRAR.exe"
    if os.path.exists(winrar_path):
        return "COPY"
        
    return "DOWNLOAD"

def perform_startup_check(root):
    status = check_unrar_status()
    if status == "OK":
        return

    if status == "COPY":
        msg = (
            "UnRAR.exe Detected in Program Files!\n\n"
            "To enable RAR/CBR support, please COPY 'UnRAR.exe' from:\n"
            "C:\\Program Files\\WinRAR\\\n\n"
            "And PASTE it into this application's folder:\n"
            f"{APP_DIR}\n\n"
            "Then restart the application."
        )
        messagebox.showinfo("Setup Required", msg, parent=root)
    
    elif status == "DOWNLOAD":
        url = "https://www.win-rar.com/start.html?&L=0"
        msg = (
            "WinRAR Not Installed / UnRAR Missing!\n\n"
            "To support RAR/CBR files, you need WinRAR.\n"
            "Would you like to open the download page now?"
        )
        if messagebox.askyesno("Missing Component", msg, icon="warning", parent=root):
            webbrowser.open(url)

# --- GUI Class ---
class ConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Run Check
        # Run Smart Check
        self.after(500, lambda: perform_startup_check(self))

        self.style = ttk.Style(theme=SETTINGS['Theme'])
        self.title("Image2PDF Professional")
        self.geometry("700x700")
        
        # Ensure icon support if we add one later
        # try: self.iconbitmap(os.path.join(APP_DIR, 'icon.ico'))
        # except: pass
        
        self.source_path = tk.StringVar()
        self.output_path = tk.StringVar(value=SETTINGS['Output_Path'])
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar()
        
        self.parallel_var = tk.BooleanVar(value=SETTINGS['Parallel_Processing'])
        self.quality_var = tk.IntVar(value=SETTINGS['PDF_Quality'])
        self.threads_var = tk.IntVar(value=SETTINGS['Thread_Count'])
        self.rar_var = tk.BooleanVar(value=SETTINGS['Enable_RAR'])
        self.landscape_var = tk.StringVar(value=SETTINGS['Landscape_Mode'])
        self.theme_var = tk.StringVar(value=SETTINGS['Theme'])
        self.delete_var = tk.BooleanVar(value=SETTINGS['Delete_Source'])
        
        self.is_running = False
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        ttk.Label(main_frame, text="Image to PDF Converter", font=("Segoe UI", 18, "bold"), bootstyle="primary").pack(pady=(0, 20))
        
        # Source
        src_frame = ttk.Labelframe(main_frame, text="Source Folder/Archive", padding=10, bootstyle="info")
        src_frame.pack(fill=X, pady=5)
        ttk.Entry(src_frame, textvariable=self.source_path).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        btn_src = ttk.Button(src_frame, text="Browse", command=self.browse_source, bootstyle="secondary-outline")
        btn_src.pack(side=LEFT)
        ToolTip(btn_src, text="Select the folder containing images or subfolders/archives")

        # Output
        out_frame = ttk.Labelframe(main_frame, text="Output Folder (Optional)", padding=10, bootstyle="info")
        out_frame.pack(fill=X, pady=5)
        ttk.Entry(out_frame, textvariable=self.output_path).pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        btn_out = ttk.Button(out_frame, text="Browse", command=self.browse_output, bootstyle="secondary-outline")
        btn_out.pack(side=LEFT)
        ToolTip(btn_out, text="PDFs will be saved here. Leave empty to save in source folder.")
        
        # Config
        opt_frame = ttk.Labelframe(main_frame, text="Configuration", padding=10, bootstyle="warning")
        opt_frame.pack(fill=X, pady=10)
        
        # Row 1
        r1 = ttk.Frame(opt_frame); r1.pack(fill=X, pady=5)
        cb_par = ttk.Checkbutton(r1, text="Parallel Processing", variable=self.parallel_var, bootstyle="round-toggle", command=self.toggle_threads)
        cb_par.pack(side=LEFT)
        ToolTip(cb_par, text="Faster conversion using multiple CPU cores")
        
        ttk.Label(r1, text="Threads:").pack(side=LEFT, padx=(15, 5))
        self.spin_threads = ttk.Spinbox(r1, from_=1, to=32, textvariable=self.threads_var, width=5)
        self.spin_threads.pack(side=LEFT)
        
        ttk.Label(r1, text="Theme:").pack(side=LEFT, padx=(20, 5))
        cb_theme = ttk.Combobox(r1, textvariable=self.theme_var, values=self.style.theme_names(), width=10, state="readonly")
        cb_theme.pack(side=LEFT)
        cb_theme.bind("<<ComboboxSelected>>", self.change_theme)

        # Row 2
        r2 = ttk.Frame(opt_frame); r2.pack(fill=X, pady=5)
        ttk.Label(r2, text="Quality (1-100):").pack(side=LEFT, padx=(0, 5))
        sp_qual = ttk.Spinbox(r2, from_=1, to=100, textvariable=self.quality_var, width=5)
        sp_qual.pack(side=LEFT)
        ToolTip(sp_qual, text="Set PDF Quality (1-100). 100 is Lossless.")
        
        ttk.Label(r2, text="Landscape Mode:").pack(side=LEFT, padx=(20, 5))
        modes = ['none', 'letterbox', 'split', 'rotate']
        cb_land = ttk.Combobox(r2, textvariable=self.landscape_var, values=modes, width=10, state="readonly")
        cb_land.pack(side=LEFT)
        ToolTip(cb_land, text="None: Keep Original\nLetterbox: Add White Bars\nSplit: Cut for Manga\nRotate: Turn 90 deg")
        
        cb_rar = ttk.Checkbutton(r2, text="RAR/CBR Support", variable=self.rar_var, bootstyle="square-toggle")
        cb_rar.pack(side=LEFT, padx=(20, 0))
        ToolTip(cb_rar, text="Enable support for .rar and .cbr files. Requires UnRAR.exe.")

        # Row 3 (Delete Safety)
        r3 = ttk.Frame(opt_frame); r3.pack(fill=X, pady=(10, 5))
        cb_del = ttk.Checkbutton(r3, text="Delete Source Files After Conversion", variable=self.delete_var, bootstyle="danger-round-toggle")
        cb_del.pack(side=LEFT)
        ToolTip(cb_del, text="WARNING: This deletes the original folder/archive after successful PDF creation.")

        btn_box = ttk.Frame(opt_frame); btn_box.pack(fill=X, pady=(10,0))
        btn_log = ttk.Button(btn_box, text="View Logs", command=self.view_logs, bootstyle="danger-link")
        btn_log.pack(side=LEFT)
        ToolTip(btn_log, text="Click to see error logs (e.g. broken zips, missing files).")
        ttk.Button(btn_box, text="Save Settings", command=self.save_settings, bootstyle="success-outline").pack(side=RIGHT)

        # Progress
        self.pbar = ttk.Floodgauge(main_frame, variable=self.progress_var, maximum=100, bootstyle="success", mask="{}%")
        self.pbar.pack(fill=X, pady=20)
        self.lbl_status = ttk.Label(main_frame, textvariable=self.status_var, font=("Segoe UI", 9))
        self.lbl_status.pack()
        
        # Act
        act_frame = ttk.Frame(main_frame); act_frame.pack(pady=10)
        self.btn_run = ttk.Button(act_frame, text="START CONVERSION", command=self.start_conversion, bootstyle="primary-lg", width=20)
        self.btn_run.pack(side=LEFT, padx=10)
        ttk.Button(act_frame, text="EXIT", command=self.destroy, bootstyle="secondary", width=10).pack(side=LEFT, padx=10)
        
        self.toggle_threads()

    def toggle_threads(self):
        if self.parallel_var.get(): self.spin_threads.configure(state='normal')
        else: self.spin_threads.configure(state='disabled')
        
    def change_theme(self, event):
        self.style.theme_use(self.theme_var.get())

    def browse_source(self):
        p = filedialog.askdirectory(); 
        if p: self.source_path.set(p)
        
    def browse_output(self):
        p = filedialog.askdirectory(); 
        if p: self.output_path.set(p)
        
    def viewer_window(self, content):
        top = ttk.Toplevel(self)
        top.title("Log Viewer")
        top.geometry("600x400")
        st = scrolledtext.ScrolledText(top, width=80, height=20)
        st.pack(fill=BOTH, expand=True)
        st.insert(tk.END, content)
        st.configure(state='disabled')

    def view_logs(self):
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r') as f: content = f.read()
            if not content: content = "No errors logged."
        else: content = "Log file not found."
        self.viewer_window(content)
        
    def save_settings(self):
        SETTINGS.update({
            'Parallel_Processing': self.parallel_var.get(),
            'Thread_Count': self.threads_var.get(),
            'PDF_Quality': self.quality_var.get(),
            'Enable_RAR': self.rar_var.get(),
            'Theme': self.theme_var.get(),
            'Landscape_Mode': self.landscape_var.get(),
            'Output_Path': self.output_path.get(),
            'Delete_Source': self.delete_var.get()
        })
        save_config(SETTINGS)
        messagebox.showinfo("Saved", "Settings saved successfully!", parent=self)

    def start_conversion(self):
        if self.is_running: return
        path = self.source_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Invalid Source Path")
            return
            
        SETTINGS.update({
            'Output_Path': self.output_path.get(),
            'Landscape_Mode': self.landscape_var.get(),
            'Thread_Count': self.threads_var.get(),
            'PDF_Quality': self.quality_var.get(),
            'Delete_Source': self.delete_var.get()
        })
        
        self.is_running = True
        self.btn_run.configure(state='disabled')
        self.status_var.set("Scanning...")
        self.progress_var.set(0)
        
        threading.Thread(target=self.run_logic, args=(path,), daemon=True).start()

    def run_logic(self, source_path):
        global ABORT_REQUESTED; ABORT_REQUESTED = False
        try:
            items = os.listdir(source_path)
            work_items = []
            for item in items:
                item_path = os.path.join(source_path, item)
                ext = os.path.splitext(item.lower())[1]
                pdf_path = get_pdf_path(item_path)
                
                if os.path.isdir(item_path):
                     work_items.append({'type': 'folder', 'path': item_path, 'pdf_path': pdf_path})
                elif ext in ['.zip', '.cbz']:
                     work_items.append({'type': 'archive', 'path': item_path, 'pdf_path': pdf_path})
                elif ext in ['.rar', '.cbr'] and SETTINGS['Enable_RAR']:
                     work_items.append({'type': 'archive', 'path': item_path, 'pdf_path': pdf_path})
            
            total = len(work_items)
            if total == 0:
                self.after(0, lambda: messagebox.showinfo("Info", "No items found!"))
                self.is_running = False
                self.after(0, lambda: self.btn_run.configure(state='normal'))
                return
            
            self.status_var.set(f"Found {total} items. Starting...")
            completed = 0
            
            if self.parallel_var.get():
                max_workers = SETTINGS['Thread_Count'] if SETTINGS['Thread_Count'] > 0 else None
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(worker_task, item): item for item in work_items}
                    for future in concurrent.futures.as_completed(futures):
                        completed += 1
                        progress = (completed / total) * 100
                        self.after(0, lambda p=progress: self.progress_var.set(p))
                        self.after(0, lambda c=completed, t=total: self.status_var.set(f"Processed {c}/{t}"))
            else:
                 for item in work_items:
                     if ABORT_REQUESTED: break
                     self.after(0, lambda n=os.path.basename(item['path']): self.status_var.set(f"Converting {n}"))
                     worker_task(item)
                     completed += 1
                     progress = (completed / total) * 100
                     self.after(0, lambda p=progress: self.progress_var.set(p))

            self.after(0, lambda: self.show_success(completed))
            
        except Exception as e:
            logging.error(f"GUI Error: {e}")
            self.after(0, lambda: messagebox.showerror("Error", f"Error: {e}"))
        finally:
            self.is_running = False
            self.after(0, lambda: self.btn_run.configure(state='normal'))
            self.after(0, lambda: self.status_var.set("Ready"))

    def show_success(self, count):
        messagebox.showinfo("Done", f"Processed {count} items.")
        if winsound: winsound.MessageBeep()
        try:
            notification.notify(
                title='Image2PDF Finished',
                message=f'Successfully converted {count} items.',
                app_name='Image2PDF',
                timeout=5
            )
        except: pass

def main():
    if len(sys.argv) > 1:
        # CLI Mode
        # Check UnRAR first
        # Check UnRAR first
        unrar_status = check_unrar_status()
        if unrar_status != "OK":
             print(f"WARNING: UnRAR status is {unrar_status}. RAR/CBR files may be skipped.\n(Check GUI mode for setup instructions)")
             
        print("CLI Mode Active.")
        source_path = sys.argv[1].strip('"').strip("'")
        if not os.path.exists(source_path): print("Path not found"); return
        
        items = os.listdir(source_path)
        work_items = []
        for item in items:
            item_path = os.path.join(source_path, item)
            ext = os.path.splitext(item.lower())[1]
            pdf_path = get_pdf_path(item_path)
            if os.path.isdir(item_path): work_items.append({'type': 'folder', 'path': item_path, 'pdf_path': pdf_path})
            elif ext in ['.zip', '.cbz']: work_items.append({'type': 'archive', 'path': item_path, 'pdf_path': pdf_path})
            elif ext in ['.rar', '.cbr'] and SETTINGS['Enable_RAR']: work_items.append({'type': 'archive', 'path': item_path, 'pdf_path': pdf_path})
        
        if SETTINGS['Parallel_Processing']:
             with tqdm(total=len(work_items)) as pbar:
                 with concurrent.futures.ThreadPoolExecutor(max_workers=SETTINGS['Thread_Count'] or None) as ex:
                     futures = [ex.submit(worker_task, i) for i in work_items]
                     for f in concurrent.futures.as_completed(futures): pbar.update(1)
        else:
            for i in tqdm(work_items): worker_task(i)
        
        try: notification.notify(title="Image2PDF", message="Batch Complete")
        except: pass
    else:
        if SETTINGS['Enable_GUI']:
            app = ConverterGUI()
            app.mainloop()

if __name__ == "__main__":
    main()

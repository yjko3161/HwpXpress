import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import sys

try:
    import win32com.client
except ImportError:
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showerror("Error", "Required modules not found.\nPlease run the application using 'run.bat' or activate the virtual environment first.\n\nError: No module named 'win32com'")
    sys.exit(1)

import threading
import os
from converter import HwpConverter
from logger_config import setup_logger

class HwpXpressApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HwpXpress - HWP to HWPX Converter")
        self.root.geometry("600x500")

        self.setup_ui()
        self.logger = setup_logger(text_widget=self.log_area)
        self.converter = HwpConverter(logger=self.logger)
        self.is_converting = False

    def setup_ui(self):
        # Frame for controls
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)

        # File Selection
        self.path_var = tk.StringVar()
        ttk.Label(control_frame, text="Target Path (File or Folder):").pack(anchor=tk.W)
        
        path_entry_frame = ttk.Frame(control_frame)
        path_entry_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(path_entry_frame, textvariable=self.path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(path_entry_frame, text="Browse File", command=self.browse_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(path_entry_frame, text="Browse Folder", command=self.browse_folder).pack(side=tk.LEFT, padx=2)

        # Action Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.convert_btn = ttk.Button(button_frame, text="Start Conversion", command=self.start_conversion)
        self.convert_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Progress Bar
        progress_frame = ttk.Frame(self.root, padding="0 5")
        progress_frame.pack(fill=tk.X, padx=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)
        
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.pack(anchor=tk.W)

        # Log Area
        log_frame = ttk.LabelFrame(self.root, text="Logs", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, state='disabled', height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("HWP files", "*.hwp")])
        if filename:
            self.path_var.set(filename)

    def browse_folder(self):
        foldername = filedialog.askdirectory()
        if foldername:
            self.path_var.set(foldername)

    def update_progress(self, current, total):
        percentage = (current / total) * 100
        self.progress_var.set(percentage)
        self.status_label.config(text=f"Processing: {current}/{total} ({percentage:.1f}%)")
        self.root.update_idletasks() # Force UI update

    def start_conversion(self):
        if self.is_converting:
            return

        target_path = self.path_var.get()
        if not target_path:
            messagebox.showwarning("Warning", "Please select a file or folder first.")
            return

        self.is_converting = True
        self.convert_btn.config(state='disabled')
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, f"\n{'='*40}\nStarting new conversion...\n{'='*40}\n")
        self.log_area.configure(state='disabled')
        self.log_area.see(tk.END)
        self.progress_var.set(0)
        self.status_label.config(text="Starting...")
        
        thread = threading.Thread(target=self.run_conversion_process, args=(target_path,))
        thread.start()

    def run_conversion_process(self, target_path):
        try:
            if os.path.isfile(target_path):
                self.update_progress(0, 1)
                self.converter.convert_to_hwpx(target_path)
                self.update_progress(1, 1)
            elif os.path.isdir(target_path):
                # Pass the update_progress method as callback
                # Since update_progress interacts with GUI, it should technically be scheduled
                # But tkinter usually handles simple var updates from threads loosely or we can use root.after
                # For safety/correctness, let's wrap it slightly if needed, but direct call often works for simple DoubleVar.
                # A safer way:
                def safe_callback(c, t):
                    self.root.after(0, lambda: self.update_progress(c, t))
                
                self.converter.convert_directory(target_path, progress_callback=safe_callback)
            else:
                self.logger.error(f"Invalid path: {target_path}")
        except Exception as e:
            self.logger.error(f"Unexpected error: {e}")
        finally:
            self.is_converting = False
            self.root.after(0, self.enable_button)

    def enable_button(self):
        self.convert_btn.config(state='normal')
        self.status_label.config(text="Conversion Complete")
        messagebox.showinfo("Complete", "Conversion process finished.")

if __name__ == "__main__":
    root = tk.Tk()
    app = HwpXpressApp(root)
    root.mainloop()

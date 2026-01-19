import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
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
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state='disabled')
        
        thread = threading.Thread(target=self.run_conversion_process, args=(target_path,))
        thread.start()

    def run_conversion_process(self, target_path):
        try:
            if os.path.isfile(target_path):
                self.converter.convert_to_hwpx(target_path)
            elif os.path.isdir(target_path):
                self.converter.convert_directory(target_path)
            else:
                self.logger.error(f"Invalid path: {target_path}")
        except Exception as e:
            self.logger.error(f"Unexpected error: {e}")
        finally:
            self.is_converting = False
            self.root.after(0, self.enable_button)

    def enable_button(self):
        self.convert_btn.config(state='normal')
        messagebox.showinfo("Complete", "Conversion process finished.")

if __name__ == "__main__":
    root = tk.Tk()
    app = HwpXpressApp(root)
    root.mainloop()

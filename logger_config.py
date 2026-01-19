import logging
import os
import sys

# Create a custom logging handler that can be connected to the GUI
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see('end')
        # This needs to be thread-safe for GUI updates, usually handle via after code in main
        # But for simplicity in tkinter we often can call directly if not too frequent or use proper event queue.
        # For now, we will assume this is attached later or handled by main app.
        try:
             self.text_widget.after(0, append)
        except Exception:
            # Fallback if widget not ready or not tkinter
            pass

def setup_logger(name="HwpXpress", log_file="app.log", text_widget=None):
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # File Handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # GUI Handler (Optional)
    if text_widget:
        gui_handler = TextHandler(text_widget)
        gui_handler.setLevel(logging.INFO)
        gui_handler.setFormatter(formatter)
        logger.addHandler(gui_handler)

    return logger

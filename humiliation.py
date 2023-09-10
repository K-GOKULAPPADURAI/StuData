from humiliation1 import *
from humiliation2 import *

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import shutil
from PIL import Image, ImageTk
class SplashScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Splash Screen")

        # Create a Label widget to display the splash image
        splash_image = Image.open("backgrounds/logo.png")  # Replace with your splash image file path
        splash_image = splash_image.resize((1100, 600), Image.Resampling.LANCZOS)
        self.splash_photo = ImageTk.PhotoImage(splash_image)
        splash_label = tk.Label(self.root, image=self.splash_photo)
        splash_label.pack()

        # Set a timer to close the splash screen after a few seconds
        self.root.after(3000, self.close_splash)

    def close_splash(self):
        self.root.destroy()
        self.open_main_window()

    def open_main_window(self):
        root = tk.Tk()
        app = StudentDetailsApp(root)


class StudentDetailsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Details Entry")
        self.root.geometry('1100x600')
                
        frame = tk.Frame(self.root)
        frame.pack()

        new_button = tk.Button(frame, text="NEW", command=self.open_new_window, font=("Helvetica", 12))
        new_button.pack(side="left", padx=10, pady=10)

        edit_button = tk.Button(frame, text="EDIT", command=self.open_edit_window, font=("Helvetica", 12))
        edit_button.pack(side="left", padx=10, pady=10)

        
        upload_button = tk.Button(frame, text="FILE UPLOAD", command=self.file_upload, font=("Helvetica", 12))
        upload_button.pack(side="left", padx=10, pady=10)

        export_button = tk.Button(frame, text="FILE EXPORT", command=self.file_export, font=("Helvetica", 12))
        export_button.pack(side="left", padx=10, pady=10)
        
 

    def open_new_window(self):
        new_window = tk.Toplevel(self.root)
        new_window.title("New Student Entry")
        NewStudentEntryApp(new_window)

    def open_edit_window(self):
        # Get the selected student's data (replace this with your data retrieval logi

        # Create the edit window using the EditStudentDetails class
        edit_window = tk.Toplevel(self.root)
        
    # Initialize and display the EditStudentApp in the editing window
        app = EditStudentApp(edit_window)

    def file_upload(self):
        # Add code for the File Upload functionality here.
        pass

    def file_export(self):
        # Add code for the File Export functionality here.
        pass
 



if __name__ == "__main__":
    root = tk.Tk()
    splash = SplashScreen(root)
    root.mainloop()

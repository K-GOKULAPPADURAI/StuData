import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl

class StudentDetailsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Details Entry")
        
        main_frame = tk.Frame(self.root)
        main_frame.pack()

        new_button = tk.Button(main_frame, text="NEW", command=self.open_new_window, font=("Helvetica", 12))
        new_button.pack(side="left", padx=10, pady=10)

        edit_button = tk.Button(main_frame, text="EDIT", command=self.open_edit_window, font=("Helvetica", 12))
        edit_button.pack(side="left", padx=10, pady=10)

        upload_button = tk.Button(main_frame, text="FILE UPLOAD", command=self.file_upload, font=("Helvetica", 12))
        upload_button.pack(side="left", padx=10, pady=10)

        export_button = tk.Button(main_frame, text="FILE EXPORT", command=self.file_export, font=("Helvetica", 12))
        export_button.pack(side="left", padx=10, pady=10)

    def open_new_window(self):
        new_window = tk.Toplevel(self.root)
        new_window.title("New Student Entry")
        NewStudentEntryApp(new_window)

    def open_edit_window(self):
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Student Details")
        # Add code for the Edit window here.

    def file_upload(self):
        # Add code for the File Upload functionality here.
        pass

    def file_export(self):
        # Add code for the File Export functionality here.
        pass

class NewStudentEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("New Student Entry")
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        labels = [
            "STUDENT BASIC DETAILS", "FAMILY DETAIL", "ACADEMIC DETAIL", 
            "SCHOLARSHIP DETAILS", "CERTIFICATE DETAILS", "BANK DETAILS"
        ]

        for label_text in labels:
            frame = tk.Frame(self.notebook)
            self.notebook.add(frame, text=label_text)

            if label_text == "STUDENT BASIC DETAILS":
                self.create_student_basic_details(frame)
            elif label_text == "FAMILY DETAIL":
                self.create_family_details(frame)
            # ... other pages ...

    # Rest of the NewStudentEntryApp class...

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentDetailsApp(root)
    root.mainloop()

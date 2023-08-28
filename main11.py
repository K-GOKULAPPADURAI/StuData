import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl

class StudentDetailsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Details Entry")
        self.root.configure(bg="light blue")

        main_frame = tk.Frame(self.root, bg="light blue")
        
        #main_frame = tk.Frame(self.root)
        main_frame.pack()

        # Load and display the background image
        bg_image = tk.PhotoImage(file=)  # Replace with your image file path
        canvas.create_image(0, 0, image=bg_image, anchor="nw")

        # Now you can add your buttons and other widgets on top of the Canvas as needed
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
        self.root.title("Student Details Entry")
        self.create_main_page()

    def create_main_page(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        # Set a background color for the main page
        main_frame.configure(bg="light blue")  # You can change the color code

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        labels = [
            "STUDENT BASIC DETAILS", "FAMILY DETAIL", "ACADEMIC DETAIL", 
            "SCHOLARSHIP DETAILS", "CERTIFICATE DETAILS", "BANK DETAILS"
        ]

        for label_text in labels:
            frame = tk.Frame(self.notebook, bg="#E5E5E5")  # Set a background color for the tab
            self.notebook.add(frame, text=label_text)

            if label_text == "STUDENT BASIC DETAILS":
                self.create_student_basic_details(frame)
            elif label_text == "FAMILY DETAIL":
                self.create_family_details(frame)
            elif label_text == "ACADEMIC DETAIL":
                self.create_academic_details(frame)
            elif label_text == "SCHOLARSHIP DETAILS":
                self.create_scholarship_details(frame)
            elif label_text == "CERTIFICATE DETAILS":
                self.create_certificate_details(frame)
            elif label_text == "BANK DETAILS":
                self.create_bank_details(frame)

        self.notebook.select(0)

        next_button = tk.Button(main_frame, text="Next", command=self.move_to_next_tab, font=("Helvetica", 12))
        next_button.pack(side="bottom", padx=10, pady=10, anchor="e")

    def move_to_next_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        next_tab = (current_tab + 1) % self.notebook.index(tk.END)
        self.notebook.select(next_tab)

    def create_student_basic_details(self, window):
        labels = [
            "APPLICATION NO:", "DATE OF APPLICATION:", "STUDENT NAME:", "AADHAR NUMBER:", "DOB:", "GENDER:",
            "DEPARTMENT:", "QUOTA:", "COMMUNITY:", "CASTE:", "RELIGION:", "MOTHER TONGUE:", "BLOOD GROUP:",
            "MEDIUM OF SCHOOL:", "MARITAL STATUS:", "FATHER NAME:", "MOTHER NAME:", "GUARDIAN NAME:",
            "STUDENT CELL NUMBER:","STUDENT PHOTO:"
        ]

        gender_options = ["Male", "Female"]
        department_options = ["CSE", "MECH", "CIVIL", "EEE", "ECE", "AI/DS"]
        religion_options = ["HINDU", "MUSLIM", "CHRISTIAN", "OTHERS"]
        language_options = ["TAMIL", "ENGLISH", "HINDI", "OTHERS"]
        blood_group_options = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"]
        school_medium_options = ["TAMIL", "ENGLISH"]
        marital_status_options = ["Single", "Married"]

        self.student_basic_data = {}

        for idx, label_text in enumerate(labels):
            if idx < len(labels) // 2:
                column = 0
                row = idx
            else:
                column = 2
                row = idx - len(labels) // 2

            label = tk.Label(window, text=label_text)
            label.grid(column=column, row=row, padx=10, pady=5, sticky="w")

            if label_text == "GENDER:":
                gender_var = tk.StringVar(value="Male")
                gender_combobox = ttk.Combobox(window, values=gender_options, textvariable=gender_var)
                gender_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = gender_var
            elif label_text == "DEPARTMENT:":
                department_var = tk.StringVar(value="CSE")
                department_combobox = ttk.Combobox(window, values=department_options, textvariable=department_var)
                department_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = department_var
            elif label_text == "RELIGION:":
                religion_var = tk.StringVar(value="HINDU")
                religion_combobox = ttk.Combobox(window, values=religion_options, textvariable=religion_var)
                religion_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = religion_var
            elif label_text == "MOTHER TONGUE:":
                language_var = tk.StringVar(value="TAMIL")
                language_combobox = ttk.Combobox(window, values=language_options, textvariable=language_var)
                language_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = language_var
            elif label_text == "BLOOD GROUP:":
                blood_group_var = tk.StringVar(value="A+")
                blood_group_combobox = ttk.Combobox(window, values=blood_group_options, textvariable=blood_group_var)
                blood_group_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = blood_group_var
            elif label_text == "MEDIUM OF SCHOOL:":
                medium_var = tk.StringVar(value="TAMIL")
                medium_combobox = ttk.Combobox(window, values=school_medium_options, textvariable=medium_var)
                medium_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = medium_var
            elif label_text == "MARITAL STATUS:":
                marital_var = tk.StringVar(value="Single")
                marital_combobox = ttk.Combobox(window, values=marital_status_options, textvariable=marital_var)
                marital_combobox.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = marital_var
            elif label_text == "STUDENT PHOTO:":  # Handle student photo input
                #photo_entry = tk.Entry(window)
                #photo_entry.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                upload_photo_button = tk.Button(window, text="Upload Photo", command=self.upload_photo)
                upload_photo_button.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")

                # Add a button to open a file dialog for selecting a photo
##                photo_button = tk.Button(window, text="Browse", command=self.browse_photo)
##                photo_button.grid(column=column + 2, row=row, padx=10, pady=5, sticky="w")

            else:
                entry = tk.Entry(window)
                entry.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = entry
    def create_family_details(self, window):
        labels = [
            "ANNUAL INCOME:", "PARENT OCCUPATION:", "JOB:", "PRESENT ADDRESS:", "PINCODE:",
            "DISTRICT:", "RESIDENCE:", "PARENT CELL NUMBER:"
        ]

        self.family_data = {}

        for label_text in labels:
            label = tk.Label(window, text=label_text)
            label.pack(side="left", padx=10, pady=5, anchor="w")

            if label_text == "PARENT OCCUPATION:":
                entry = tk.Entry(window)
                entry.pack(side="left", padx=10, pady=5, anchor="w")
                self.family_data[label_text] = entry
            elif label_text == "JOB:":
                job_var = tk.StringVar(value="Government Job")
                job_combobox = ttk.Combobox(window, values=["Government Job", "Private Job"], textvariable=job_var)
                job_combobox.pack(side="left", padx=10, pady=5, anchor="w")
                self.family_data[label_text] = job_var
            else:
                entry = tk.Entry(window)
                entry.pack(side="left", padx=10, pady=5, anchor="w")
                self.family_data[label_text] = entry
       
    def upload_photo(self):
        # Get the student's name from the entry field
        student_name = self.student_basic_data["STUDENT NAME:"].get()

        # Ensure a valid student name is provided
        if student_name:
            # Use the student's name as the file name
            file_name = f"{student_name}.jpg"

            # Use filedialog to open a file dialog for photo selection
            file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif")])

            if file_path:
                # Check if the "photos" directory exists, create it if not
                if not os.path.exists("photos"):
                    os.mkdir("photos")

                # Save the photo with the student's name as the file name in the "photos" directory
                photo_save_path = os.path.join("photos", file_name)
                shutil.copyfile(file_path, photo_save_path)
                
                messagebox.showinfo("Success", "Photo uploaded and saved successfully!")
        else:
            messagebox.showwarning("Warning", "Please enter the student's name before uploading a photo.")


    def create_academic_details(self, window):
        labels = ["School Name:", "Location:", "Board:", "Year of Passing:", "Percentage of Marks:"]
        grades = ["10th", "11th", "12th"]

        self.academic_data = []

        matrix_rows = 2
        matrix_cols = 2

        for row in range(matrix_rows):
            for col in range(matrix_cols):
                grade_idx = row * matrix_cols + col
                if grade_idx < len(grades):
                    label_text = f"{grades[grade_idx]} Grade"

                    frame = tk.Frame(window, bg="#E5E5E5")  # Set a background color for the grade frame
                    frame.pack(side="left", padx=10, pady=5, anchor="w")

                    label = tk.Label(frame, text=label_text)
                    label.pack(side="top", padx=(0, 5), anchor="w")

                    academic_section = {}
                    for sub_label_text in labels:
                        sub_label = tk.Label(frame, text=f"{sub_label_text}:")
                        sub_label.pack(side="top", padx=(0, 5), anchor="w")

                        entry = tk.Entry(frame)
                        entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")
                        academic_section[sub_label_text] = entry

                    self.academic_data.append(academic_section)

        emis_label = tk.Label(window, text="EMIS NO:")
        emis_label.pack(side="top", padx=(10, 5), pady=5, anchor="w")
        self.emis_entry = tk.Entry(window)
        self.emis_entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")

        physics_label = tk.Label(window, text="Physics Mark:")
        physics_label.pack(side="top", padx=(10, 5), pady=5, anchor="w")
        self.physics_entry = tk.Entry(window)
        self.physics_entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")

        chemistry_label = tk.Label(window, text="Chemistry Mark:")
        chemistry_label.pack(side="top", padx=(10, 5), pady=5, anchor="w")
        self.chemistry_entry = tk.Entry(window)
        self.chemistry_entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")

        maths_label = tk.Label(window, text="Maths Mark:")
        maths_label.pack(side="top", padx=(10, 5), pady=5, anchor="w")
        self.maths_entry = tk.Entry(window)
        self.maths_entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")

        cutoff_label = tk.Label(window, text="Cut Off Mark:")
        cutoff_label.pack(side="top", padx=(10, 5), pady=5, anchor="w")
        self.cutoff_entry = tk.Entry(window)
        self.cutoff_entry.pack(side="top", padx=(0, 10), pady=2, anchor="w")

    def create_scholarship_details(self, window):
        labels = ["BC:", "MBC:", "PMMS:", "NSP:", "7.5:", "MOOVALUR:", "CKCET:", "FG:", "JD:"]

        self.scholarship_data = {}

        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.pack(side="left", padx=10, pady=5, anchor="w")
            self.scholarship_data[label_text] = var

    def create_certificate_details(self, window):
        labels = [
            "10th Original:", "11th Original:", "12th Original:", "TC:", "FG:", "JD:", "Community:",
            "Income:", "Aadhar:", "P-Aadhar:", "Family/Smart card:"
        ]

        self.certificate_data = {}

        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.pack(side="left", padx=10, pady=5, anchor="w")
            self.certificate_data[label_text] = var

    def create_bank_details(self, window):
        bank_labels = ["BANK NAME:", "BANK BRANCH:", "ACCOUNT NO:", "IFSC CODE:", "MIRC CODE:"]
        self.bank_data = {}

        for label_text in bank_labels:
            label = tk.Label(window, text=label_text)
            label.pack(side="left", padx=10, pady=5, anchor="w")

            entry = tk.Entry(window)
            entry.pack(side="left", padx=10, pady=5, anchor="w")
            self.bank_data[label_text] = entry

        save_button = tk.Button(window, text="Save", command=self.save_data)
        save_button.pack(side="right", padx=10, pady=10, anchor="e")

    def save_data(self):
        wb = openpyxl.load_workbook('student_details.xlsx')
        sheet = wb.active

        data = [
            self.student_basic_data,
            self.family_data,
            *self.academic_data,
            self.scholarship_data,
            self.certificate_data,
            self.bank_data
        ]

        row = []
        for section in data:
            if isinstance(section, dict):
                for value in section.values():
                    if isinstance(value, tk.StringVar):
                        row.append(value.get())
                    elif isinstance(value, tk.Entry):
                        row.append(value.get())  # Retrieve the value from the Entry widget
                    else:
                        row.append(value)
            else:
                row.append(section.get())

        sheet.append(row)
        wb.save('student_details.xlsx')
        wb.close()
        messagebox.showinfo("Success", "Data saved successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentDetailsApp(root)
    root.mainloop()

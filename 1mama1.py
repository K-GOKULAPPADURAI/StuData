import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl

class StudentDetailsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Details Entry")
        
        self.create_main_page()

    def create_main_page(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack()

        labels = ["STUDENT BASIC DETAILS", "FAMILY DETAIL", "ACADEMIC DETAIL", "SCHOLARSHIP DETAILS", "CERTIFICATE DETAILS", "BANK DETAILS"]

        for label_text in labels:
            button = tk.Button(main_frame, text=label_text, command=lambda label=label_text: self.open_child_window(label))
            button.pack(side="top", fill="x", padx=10, pady=5)
        
    def open_child_window(self, label):
        child_window = tk.Toplevel(self.root)
        
        if label == "STUDENT BASIC DETAILS":
            self.create_student_basic_details(child_window)
        elif label == "FAMILY DETAIL":
            self.create_family_details(child_window)
        elif label == "ACADEMIC DETAIL":
            self.create_academic_details(child_window)
        elif label == "SCHOLARSHIP DETAILS":
            self.create_scholarship_details(child_window)
        elif label == "CERTIFICATE DETAILS":
            self.create_certificate_details(child_window)
        elif label == "BANK DETAILS":
            self.create_bank_details(child_window)
            
    def create_student_basic_details(self, window):
        labels = [
            "APPLICATION NO:", "DATE OF APPLICATION:", "STUDENT NAME:", "AADHAR NUMBER:", "DOB:", "GENDER:",
            "DEPARTMENT:", "QUOTA:", "COMMUNITY:", "CASTE:", "RELIGION:", "MOTHER TONGUE:", "BLOOD GROUP:",
            "MEDIUM OF SCHOOL:", "MARITAL STATUS:", "FATHER NAME:", "MOTHER NAME:", "GUARDIAN NAME:",
            "STUDENT CELL NUMBER:"
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
            else:
                entry = tk.Entry(window)
                entry.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = entry

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window("FAMILY DETAIL"))
        next_button.grid(column=1, row=len(labels) // 2, padx=10, pady=10, sticky="e")
    def create_family_details(self, window):
            labels = ["ANNUAL INCOME:", "PARENT OCCUPATION:","JOB:","PRESENT ADDRESS:", "PINCODE:", 
                      "DISTRICT:", "RESIDENCE:", "PARENT CELL NUMBER:"]
            
            self.family_data = {}
            
            for label_text in labels:
                label = tk.Label(window, text=label_text)
                label.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                
                if label_text == "PARENT OCCUPATION:":
                    entry = tk.Entry(window)
                    entry.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                    self.family_data[label_text] = entry
                elif label_text == "JOB:":
                    job_var = tk.StringVar(value="Government Job")
                    job_combobox = ttk.Combobox(window, values=["Government Job", "Private Job"], textvariable=job_var)
                    job_combobox.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                    self.family_data[label_text] = job_var
                else:
                    entry = tk.Entry(window)
                    entry.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                    self.family_data[label_text] = entry

            next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window("ACADEMIC DETAIL"))
            next_button.grid(column=1, row=len(labels) + 1, padx=10, pady=10, sticky="e")
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
                    
                    frame = tk.Frame(window)
                    frame.grid(row=row, column=col, padx=10, pady=5, sticky="w")

                    label = tk.Label(frame, text=label_text)
                    label.grid(row=0, column=0, padx=(0, 5), sticky="w")

                    academic_section = {}
                    for sub_row in range(len(labels)):
                        sub_label = tk.Label(frame, text=f"{labels[sub_row]}:")
                        sub_label.grid(row=sub_row + 1, column=0, padx=(0, 5), sticky="w")

                        entry = tk.Entry(frame)
                        entry.grid(row=sub_row + 1, column=1, padx=(0, 10), pady=2, sticky="w")
                        academic_section[labels[sub_row]] = entry

                    self.academic_data.append(academic_section)

        emis_label = tk.Label(window, text="EMIS NO:")
        emis_label.grid(row=matrix_rows * matrix_cols, column=0, padx=(10, 5), pady=5, sticky="w")
        self.emis_entry = tk.Entry(window)
        self.emis_entry.grid(row=matrix_rows * matrix_cols, column=1, padx=(0, 10), pady=2, sticky="w")

        physics_label = tk.Label(window, text="Physics Mark:")
        physics_label.grid(row=matrix_rows * matrix_cols + 1, column=0, padx=(10, 5), pady=5, sticky="w")
        self.physics_entry = tk.Entry(window)
        self.physics_entry.grid(row=matrix_rows * matrix_cols + 1, column=1, padx=(0, 10), pady=2, sticky="w")

        chemistry_label = tk.Label(window, text="Chemistry Mark:")
        chemistry_label.grid(row=matrix_rows * matrix_cols + 2, column=0, padx=(10, 5), pady=5, sticky="w")
        self.chemistry_entry = tk.Entry(window)
        self.chemistry_entry.grid(row=matrix_rows * matrix_cols + 2, column=1, padx=(0, 10), pady=2, sticky="w")

        maths_label = tk.Label(window, text="Maths Mark:")
        maths_label.grid(row=matrix_rows * matrix_cols + 3, column=0, padx=(10, 5), pady=5, sticky="w")
        self.maths_entry = tk.Entry(window)
        self.maths_entry.grid(row=matrix_rows * matrix_cols + 3, column=1, padx=(0, 10), pady=2, sticky="w")

        cutoff_label = tk.Label(window, text="Cut Off Mark:")
        cutoff_label.grid(row=matrix_rows * matrix_cols + 4, column=0, padx=(10, 5), pady=5, sticky="w")
        self.cutoff_entry = tk.Entry(window)
        self.cutoff_entry.grid(row=matrix_rows * matrix_cols + 4, column=1, padx=(0, 10), pady=2, sticky="w")

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window("SCHOLARSHIP DETAILS"))
        next_button.grid(row=matrix_rows * matrix_cols + 5, column=1, padx=10, pady=10, sticky="e")
    def create_scholarship_details(self, window):
        labels = ["BC:", "MBC:", "PMMS:", "NSP:", "7.5:", "MOOVALUR:", "CKCET:", "FG:","JD:"]
        
        self.scholarship_data = {}
        
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.scholarship_data[label_text] = var

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window("CERTIFICATE DETAILS"))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    
    def create_certificate_details(self, window):
        labels = ["10th Original:", "11th Original:", "12th Original:", "TC:", "FG:", "JD:", "Community:",
                   "Income:", "Aadhar:", "P-Aadhar:", "Family/Smart card:"]

        self.certificate_data = {}
            
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.certificate_data[label_text] = var

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window("BANK DETAILS"))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def create_bank_details(self, window):
        bank_labels = ["BANK NAME:", "BANK BRANCH:", "ACCOUNT NO:", "IFSC CODE:", "MIRC CODE:"]
        self.bank_data = {}

        for label_text in bank_labels:
            label = tk.Label(window, text=label_text)
            label.grid(column=0, row=bank_labels.index(label_text), padx=10, pady=5, sticky="w")

            entry = tk.Entry(window)
            entry.grid(column=1, row=bank_labels.index(label_text), padx=10, pady=5, sticky="w")
            self.bank_data[label_text] = entry

        save_button = tk.Button(window, text="Save", command=self.save_data)
        save_button.grid(column=1, row=len(bank_labels), padx=10, pady=10, sticky="e")
    def create_search_student_window(self, window):
        window.title("Search Student")
        
        # Create and configure GUI elements for searching students
        search_label = tk.Label(window, text="Enter Student ID:")
        search_label.pack(pady=10)
        
        search_entry = tk.Entry(window)
        search_entry.pack(pady=5)
        
        search_button = tk.Button(window, text="Search", command=lambda: self.search_student(search_entry.get()))
        search_button.pack(pady=10)
        
        # Display student information here (e.g., name, details, and photo)
        student_info_label = tk.Label(window, text="Student Info:")
        student_info_label.pack(pady=10)
        
        self.student_info_text = tk.Text(window, height=10, width=50)
        self.student_info_text.pack()

        # Display student photo here
        self.student_photo = tk.Label(window)
        self.student_photo.pack()

    def create_upload_excel_window(self, window):
        window.title("Upload Excel")
        
        # Create and configure GUI elements for uploading Excel file
        upload_label = tk.Label(window, text="Select Excel File:")
        upload_label.pack(pady=10)
        
        self.uploaded_file_path = tk.StringVar()
        
        upload_button = tk.Button(window, text="Upload", command=self.upload_excel)
        upload_button.pack(pady=5)
        
        append_button = tk.Button(window, text="Append to Current Data", command=self.append_to_excel)
        append_button.pack(pady=10)

    def create_export_excel_window(self, window):
        window.title("Export Excel")
        
        # Create and configure GUI elements for exporting Excel file
        export_label = tk.Label(window, text="Select Criteria:")
        export_label.pack(pady=10)
        
        self.export_criteria_var = tk.StringVar()
        export_entry = tk.Entry(window, textvariable=self.export_criteria_var)
        export_entry.pack(pady=5)
        
        export_button = tk.Button(window, text="Export Data", command=self.export_excel)
        export_button.pack(pady=10)

    # Implement the functionality for these functions below:

    def search_student(self, student_id):
        # Placeholder: Implement the logic to search for a student and display their information
        student_info = f"Student ID: {student_id}\n"  # Replace with actual student data
        self.student_info_text.delete("1.0", tk.END)  # Clear existing text
        self.student_info_text.insert(tk.END, student_info)
        
        # Placeholder: Implement logic to display the student's photo (you need to load and display the image)
        # Example: Load a sample image for testing
        image = Image.open("sample_student_photo.jpg")
        photo = ImageTk.PhotoImage(image)
        self.student_photo.config(image=photo)
        self.student_photo.image = photo  # Keep a reference to avoid garbage collection

    def upload_excel(self):
        # Placeholder: Implement the logic to upload an Excel file
        file_path = tk.filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.uploaded_file_path.set(file_path)

    def append_to_excel(self):
        # Placeholder: Implement the logic to append data from the uploaded Excel file to the current data
        uploaded_file_path = self.uploaded_file_path.get()
        if uploaded_file_path:
            # Load and append data from the uploaded Excel file
            # Example: Append data using openpyxl or pandas library
            pass

    def export_excel(self):
        criteria = self.export_criteria_var.get()
        if criteria:
            # Placeholder: Implement the logic to export data based on the specified criteria
            # Example: Export data using openpyxl or pandas library
            pass


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
                row.extend([value.get() if isinstance(value, tk.StringVar) else value for value in section.values()])
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

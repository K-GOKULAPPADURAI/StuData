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
            if idx < len(labels) // 2:  # Divide labels into two columns
                column = 0
            else:
                column = 2

            label = tk.Label(window, text=label_text)
            label.grid(column=column, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")

            if label_text == "GENDER:":
                gender_var = tk.StringVar(value="Male")
                gender_combobox = ttk.Combobox(window, values=gender_options, textvariable=gender_var)
                gender_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = gender_var

            elif label_text == "DEPARTMENT:":
                department_var = tk.StringVar(value="CSE")
                department_combobox = ttk.Combobox(window, values=department_options, textvariable=department_var)
                department_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = department_var
            elif label_text == "RELIGION:":
                religion_var = tk.StringVar(value="HINDU")
                religion_combobox = ttk.Combobox(window, values=religion_options, textvariable=religion_var)
                religion_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = religion_var
            elif label_text == "MOTHER TONGUE:":
                language_var = tk.StringVar(value="TAMIL")
                language_combobox = ttk.Combobox(window, values=language_options, textvariable=language_var)
                language_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = language_var
            elif label_text == "BLOOD GROUP:":
                blood_group_var = tk.StringVar(value="A+")
                blood_group_combobox = ttk.Combobox(window, values=blood_group_options, textvariable=blood_group_var)
                blood_group_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = blood_group_var
            elif label_text == "MEDIUM OF SCHOOL:":
                medium_var = tk.StringVar(value="TAMIL")
                medium_combobox = ttk.Combobox(window, values=school_medium_options, textvariable=medium_var)
                medium_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = medium_var
            elif label_text == "MARITAL STATUS:":
                marital_var = tk.StringVar(value="Single")
                marital_combobox = ttk.Combobox(window, values=marital_status_options, textvariable=marital_var)
                marital_combobox.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = marital_var
            else:
                entry = tk.Entry(window)
                entry.grid(column=column + 1, row=idx % (len(labels) // 2), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = entry

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(1))
        next_button.grid(column=1, row=len(labels) // 2, padx=10, pady=10, sticky="e")



    def create_family_details(self, window):
        labels = ["ANNUAL INCOME:", "PARENT OCCUPATION:", "PRESENT ADDRESS:", "PINCODE:", 
                  "DISTRICT:", "RESIDENCE:", "PARENT CELL NUMBER:"]
        
        self.family_data = {}
        
        for label_text in labels:
            label = tk.Label(window, text=label_text)
            label.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            
            if label_text == "PARENT OCCUPATION:":
                entry = tk.Entry(window)
                entry.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.family_data[label_text] = entry
            elif label_text == "RESIDENCE:":
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

        for grade in grades:
            academic_section = {}
            label_text = f"{grade} Grade"
            for row in range(len(labels)):
                label = tk.Label(window, text=f"{label_text} - {labels[row]}:")
                label.grid(column=0, row=row, padx=10, pady=5, sticky="w")

                entry = tk.Entry(window)
                entry.grid(column=1, row=row, padx=10, pady=5, sticky="w")
                academic_section[labels[row]] = entry
            self.academic_data.append(academic_section)

        physics_label = tk.Label(window, text="Physics Mark:")
        physics_label.grid(column=2, row=0, padx=10, pady=5, sticky="w")
        self.physics_entry = tk.Entry(window)
        self.physics_entry.grid(column=3, row=0, padx=10, pady=5, sticky="w")

        chemistry_label = tk.Label(window, text="Chemistry Mark:")
        chemistry_label.grid(column=2, row=1, padx=10, pady=5, sticky="w")
        self.chemistry_entry = tk.Entry(window)
        self.chemistry_entry.grid(column=3, row=1, padx=10, pady=5, sticky="w")

        maths_label = tk.Label(window, text="Maths Mark:")
        maths_label.grid(column=2, row=2, padx=10, pady=5, sticky="w")
        self.maths_entry = tk.Entry(window)
        self.maths_entry.grid(column=3, row=2, padx=10, pady=5, sticky="w")

        cutoff_label = tk.Label(window, text="Cut Off Mark:")
        cutoff_label.grid(column=2, row=3, padx=10, pady=5, sticky="w")
        self.cutoff_entry = tk.Entry(window)
        self.cutoff_entry.grid(column=3, row=3, padx=10, pady=5, sticky="w")

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(3))
        next_button.grid(column=3, row=len(labels), padx=10, pady=10, sticky="e")

    def create_scholarship_details(self, window):
        labels = ["BC:", "MBC:", "PMMS:", "NSP:", "7.5:", "MOOVALUR:", "CKCET:", "FG:","JG:"]
        
        self.scholarship_data = {}
        
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.scholarship_data[label_text] = var

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(4))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def create_certificate_details(self, window):
        labels = ["10th Original:", "11th Original:", "12th Original:", "TC:", "FG:","JG:", "Community:"
                   "Income:","Aadhar:","P-Aadhar:","Family/Smart card:" ]
        
        self.certificate_data = {}
            
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.certificate_data[label_text] = var

        save_button = tk.Button(window, text="Save", command=self.save_data)
        save_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

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

    def save_data(self):
        wb = openpyxl.load_workbook('student_details.xlsx')
        sheet = wb.active

        data = [
            self.student_basic_data,
            self.family_data,
            *self.academic_data,
            {
                "Physics Mark": self.physics_entry.get(),
                "Chemistry Mark": self.chemistry_entry.get(),
                "Maths Mark": self.maths_entry.get(),
                "Cut Off Mark": self.cutoff_entry.get()
            },
            self.scholarship_data,
            self.certificate_data,
            self.bank_data
        ]

        row = []
        for section in data:
            row.extend([value.get() if isinstance(value, tk.StringVar) else value.get() for value in section.values()])

        sheet.append(row)
        wb.save('student_details.xlsx')
        wb.close()
        messagebox.showinfo("Success", "Data saved successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentDetailsApp(root)
    root.mainloop()

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

        labels = ["STUDENT BASIC DETAILS", "FAMILY DETAIL", "ACADEMIC DETAIL", "SCHOLARSHIP DETAILS", "CERTIFICATE DETAILS"]

        for idx, label_text in enumerate(labels):
            button = tk.Button(main_frame, text=label_text, command=lambda idx=idx: self.open_child_window(idx))
            button.pack(pady=10)

    def open_child_window(self, idx):
        child_window = tk.Toplevel(self.root)
        
        if idx == 0:
            self.create_student_basic_details(child_window)
        elif idx == 1:
            self.create_family_details(child_window)
        elif idx == 2:
            self.create_academic_details(child_window)
        elif idx == 3:
            self.create_scholarship_details(child_window)
        elif idx == 4:
            self.create_certificate_details(child_window)

    def create_student_basic_details(self, window):
        labels = ["APPLICATION NO:", "DATE OF APPLICATION:", "STUDENT NAME:", "AADHAR NUMBER:", "DOB:", "GENDER:", 
                  "DEPARTMENT:", "QUOTA:", "COMMUNITY:", "CASTE:", "RELIGION:", "MOTHER TONGUE:", "BLOOD GROUP:", 
                  "MEDIUM OF SCHOOL:", "MARITAL STATUS:", "FATHER NAME:", "MOTHER NAME:", "GUARDIAN NAME:", 
                  "STUDENT CELL NUMBER:"]
        
        gender_options = ["Male", "Female"]
        department_options = ["CSE", "MECH", "CIVIL", "EEE", "ECE", "AI/DS"]
        religion_options = ["HINDU", "MUSLIM", "CHRISTIAN", "PARSI"]
        language_options = ["TAMIL", "ENGLISH", "HINDI", "ARABIC", "FRENCH", "MALAYALAM", "MARATHI"]
        blood_group_options = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"]
        school_medium_options = ["TAMIL", "ENGLISH"]
        marital_status_options = ["Single", "Married"]

        self.student_basic_data = {}
        
        for label_text in labels:
            label = tk.Label(window, text=label_text)
            label.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            
            if label_text == "GENDER:":
                gender_var = tk.StringVar(value="Male")
                for i, gender_option in enumerate(gender_options):
                    radio = tk.Radiobutton(window, text=gender_option, variable=gender_var, value=gender_option)
                    radio.grid(column=1, row=labels.index(label_text)+i, sticky="w")
                self.student_basic_data[label_text] = gender_var
            elif label_text == "DEPARTMENT:":
                department_var = tk.StringVar(value="CSE")
                department_combobox = ttk.Combobox(window, values=department_options, textvariable=department_var)
                department_combobox.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = department_var
            elif label_text == "RELIGION:":
                religion_var = tk.StringVar(value="HINDU")
                religion_combobox = ttk.Combobox(window, values=religion_options, textvariable=religion_var)
                religion_combobox.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = religion_var
            elif label_text == "MOTHER TONGUE:":
                language_var = tk.StringVar(value="TAMIL")
                language_combobox = ttk.Combobox(window, values=language_options, textvariable=language_var)
                language_combobox.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = language_var
            elif label_text == "BLOOD GROUP:":
                blood_group_var = tk.StringVar(value="A+")
                blood_group_combobox = ttk.Combobox(window, values=blood_group_options, textvariable=blood_group_var)
                blood_group_combobox.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = blood_group_var
            elif label_text == "MEDIUM OF SCHOOL:":
                school_medium_var = tk.StringVar(value="TAMIL")
                for i, medium_option in enumerate(school_medium_options):
                    radio = tk.Radiobutton(window, text=medium_option, variable=school_medium_var, value=medium_option)
                    radio.grid(column=1, row=labels.index(label_text)+i, sticky="w")
                self.student_basic_data[label_text] = school_medium_var
            elif label_text == "MARITAL STATUS:":
                marital_status_var = tk.StringVar(value="Single")
                for i, status_option in enumerate(marital_status_options):
                    radio = tk.Radiobutton(window, text=status_option, variable=marital_status_var, value=status_option)
                    radio.grid(column=1, row=labels.index(label_text)+i, sticky="w")
                self.student_basic_data[label_text] = marital_status_var
            else:
                entry = tk.Entry(window)
                entry.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = entry

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(1))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def create_family_details(self, window):
        labels = ["ANNUAL INCOME:", "PARENT OCCUPATION:", "GOVERNEMENT OR PRIVATE:", "PRESENT ADDRESS:", "PINCODE:", 
                  "DISTRICT:", "RESIDENCE:", "PARENT CELL NUMBER:"]
        
        gov_or_private_options = ["Yes", "No"]

        self.family_data = {}
        
        for label_text in labels:
            label = tk.Label(window, text=label_text)
            label.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            
            if label_text == "GOVERNEMENT OR PRIVATE:":
                gov_or_private_var = tk.StringVar(value="Yes")
                for i, option in enumerate(gov_or_private_options):
                    radio = tk.Radiobutton(window, text=option, variable=gov_or_private_var, value=option)
                    radio.grid(column=1, row=labels.index(label_text)+i, sticky="w")
                self.family_data[label_text] = gov_or_private_var
            else:
                entry = tk.Entry(window)
                entry.grid(column=1, row=labels.index(label_text), padx=10, pady=5, sticky="w")
                self.family_data[label_text] = entry

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(2))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def create_academic_details(self, window):
        labels = ["Xth School Name:", "Location:", "Board:", "Year of Passing:", "Percentage of Marks:"]
        columns = 3  # Number of columns for academic details

        self.academic_data = []
        
        for row in range(len(labels)):
            academic_section = {}
            for col in range(columns):
                label_text = labels[row]
                label = tk.Label(window, text=f"{label_text} ({col + 10}):")
                label.grid(column=col * 2, row=row, padx=10, pady=5, sticky="w")

                entry = tk.Entry(window)
                entry.grid(column=col * 2 + 1, row=row, padx=10, pady=5, sticky="w")
                academic_section[label_text] = entry
            self.academic_data.append(academic_section)

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(3))
        next_button.grid(column=columns * 2 + 1, row=len(labels), padx=10, pady=10, sticky="e")


    def create_scholarship_details(self, window):
        labels = ["BC:", "MBC:", "PMMS:", "NSP:", "7.5:", "MOOVALUR:", "CKCET:", "FC:"]
        
        self.scholarship_data = {}
        
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.scholarship_data[label_text] = var

        next_button = tk.Button(window, text="Next", command=lambda: self.open_child_window(4))
        next_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def create_certificate_details(self, window):
        labels = ["10th Original:", "11th Original:", "12th Original:", "TC:", "FC:", "Community:"]
        
        self.certificate_data = {}
        
        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.grid(column=0, row=labels.index(label_text), padx=10, pady=5, sticky="w")
            self.certificate_data[label_text] = var

        save_button = tk.Button(window, text="Save", command=self.save_data)
        save_button.grid(column=1, row=len(labels), padx=10, pady=10, sticky="e")

    def save_data(self):
        wb = openpyxl.load_workbook('student_details.xlsx')
        sheet = wb.active

        data = [
            self.student_basic_data,
            self.family_data,
            *self.academic_data,
            self.scholarship_data,
            self.certificate_data
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

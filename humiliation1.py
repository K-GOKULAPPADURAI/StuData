import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import shutil
from PIL import Image, ImageTk

class NewStudentEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Details Entry")
        self.root.geometry('1100x600')
        main_frame = tk.Frame(self.root)
        main_frame.pack()
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)    

        labels = [
            "STUDENT BASIC DETAILS", "FAMILY DETAIL", "ACADEMIC DETAIL", 
            "CERTIFICATE DETAILS","SCHOLARSHIP DETAILS",  "BANK DETAILS"
        ]

        for label_text in labels:
            frame = tk.Frame(self.notebook)
            self.notebook.add(frame, text=label_text)

            if label_text == "STUDENT BASIC DETAILS":
                self.create_student_basic_details(frame)
            elif label_text == "FAMILY DETAIL":
                self.create_family_details(frame)
            elif label_text == "ACADEMIC DETAIL":
                self.create_academic_details(frame)
            elif label_text == "CERTIFICATE DETAILS":
                self.create_certificate_details(frame)
            elif label_text == "SCHOLARSHIP DETAILS":
                self.create_scholarship_details(frame)
            elif label_text == "BANK DETAILS":
                self.create_bank_details(frame)

        self.notebook.select(0)
        next_button = tk.Button(frame, text="Next", command=self.move_to_next_tab, font=("Helvetica", 12))
        next_button.pack(side="bottom", padx=10, pady=10, anchor="e")

    def create_student_basic_details(self, window):
        background_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your background image path
        background_label = tk.Label(window, image=background_image)
        background_label.photo = background_image  # Keep a reference to the PhotoImage
        background_label.place(relwidth=1, relheight=1)

        labels = [
            "APPLICATION NO:", "DATE OF APPLICATION:", "STUDENT NAME:", "AADHAR NUMBER:", "DOB:", "GENDER:",
            "DEPARTMENT:", "QUOTA:", "COMMUNITY:", "CASTE:", "RELIGION:", "MOTHER TONGUE:", "BLOOD GROUP:",
            "MEDIUM OF SCHOOL:", "MARITAL STATUS:", "FATHER NAME:", "MOTHER NAME:", "GUARDIAN NAME:", "STUDENT PHOTO:"
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
            elif label_text == "STUDENT PHOTO:":
                upload_photo_button = tk.Button(window, text="Upload Photo", command=self.upload_photo)
                upload_photo_button.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
            else:
                entry = tk.Entry(window)
                entry.grid(column=column + 1, row=row, padx=10, pady=5, sticky="w")
                self.student_basic_data[label_text] = entry

    def move_to_next_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        next_tab = (current_tab + 1) % self.notebook.index(tk.END)
        self.notebook.select(next_tab)

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

    def create_family_details(self, window):
        # Set the background image for the window
        background_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your background image path
        background_label = tk.Label(window, image=background_image)
        background_label.photo = background_image  # Keep a reference to the PhotoImage
        background_label.place(relwidth=1, relheight=1)

        labels = ["ANNUAL INCOME:","PARENT OCCUPATION","JOB:", "PRESENT ADDRESS:", "PERMANENT ADDRESS:", "PINCODE:", "DISTRICT:",
                  "RESIDENCE:", "PARENT CELL NUMBER:", "STUDENT CELL NUMBER:", "HOSTEL:"]

        self.family_data = {}

        for idx, label_text in enumerate(labels):
            label = tk.Label(window, text=label_text)
            label.place(x=20, y=20 + idx * 30)  # Adjust the x and y coordinates as needed

            if label_text == "JOB:":
                job_var = tk.StringVar(value="Government Job")
                job_combobox = ttk.Combobox(window, values=["Government Job", "Private Job"], textvariable=job_var)
                job_combobox.place(x=200, y=20 + idx * 30)  # Adjust the x coordinate as needed
                self.family_data[label_text] = job_var
            elif label_text == "HOSTEL:":
                h_var = tk.StringVar(value="no")
                h_combobox = ttk.Combobox(window, values=["no", "yes"], textvariable=h_var)
                h_combobox.place(x=200, y=20 + idx * 30)  # Adjust the x coordinate as needed
                self.family_data[label_text] = h_var
            
            else:
                entry = tk.Entry(window)
                entry.place(x=200, y=20 + idx * 30)  # Adjust the x coordinate as needed
                self.family_data[label_text] = entry
    def create_academic_details(self, window):
        # Set the background image for the window
        background_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your background image path
        background_label = tk.Label(window, image=background_image)
        background_label.photo = background_image  # Keep a reference to the PhotoImage
        background_label.place(relwidth=1, relheight=1)

        labels = ["School Name:", "Location:", "School Type:", "Board:", "Year of Passing:", "Percentage of Marks:"]
        grades = ["10th", "11th", "12th"]

        self.academic_data = []

        for idx, grade in enumerate(grades):
            academic_section = {}  # Create a new academic_section dictionary for each grade

            label_text = f"{grade} Standard"

            frame = tk.Frame(window)
            frame.grid(row=idx // 2, column=idx % 2, padx=10, pady=10, sticky="nsew")  # Adjust the padding as needed

            label = tk.Label(frame, text=label_text)
            label.grid(row=0, column=0, padx=(0, 5), sticky="w")

            for sub_row, label_text in enumerate(labels):
                sub_label = tk.Label(frame, text=f"{label_text}:")
                sub_label.grid(row=sub_row + 1, column=0, padx=(0, 5), sticky="w")

                if label_text == "School Type:" and grade == "12th":
                    school_type_var = tk.StringVar(value="Government")
                    school_type_combobox = ttk.Combobox(frame, values=["Government", "Government Aided", "Private"], textvariable=school_type_var)
                    school_type_combobox.grid(row=sub_row + 1, column=1, padx=(0, 10), pady=2, sticky="w")
                    academic_section[label_text] = school_type_var
                elif label_text!="School Type:":
                    entry = tk.Entry(frame)
                    entry.grid(row=sub_row + 1, column=1, padx=(0, 10), pady=2, sticky="w")
                    academic_section[label_text] = entry

            # Add Physics Mark, Chemistry Mark, Maths Mark, Cut Off Mark, and EMIS NO fields to the academic section dictionary
            if grade == "12th":
                physics_label = tk.Label(frame, text="Physics Mark:")
                physics_label.grid(row=len(labels) + 1, column=0, padx=(0, 5), sticky="w")
                physics_entry = tk.Entry(frame)
                physics_entry.grid(row=len(labels) + 1, column=1, padx=(0, 10), pady=2, sticky="w")
                academic_section["Physics Mark"] = physics_entry

                chemistry_label = tk.Label(frame, text="Chemistry Mark:")
                chemistry_label.grid(row=len(labels) + 2, column=0, padx=(0, 5), sticky="w")
                chemistry_entry = tk.Entry(frame)
                chemistry_entry.grid(row=len(labels) + 2, column=1, padx=(0, 10), pady=2, sticky="w")
                academic_section["Chemistry Mark"] = chemistry_entry

                maths_label = tk.Label(frame, text="Maths Mark:")
                maths_label.grid(row=len(labels) + 3, column=0, padx=(0, 5), sticky="w")
                maths_entry = tk.Entry(frame)
                maths_entry.grid(row=len(labels) + 3, column=1, padx=(0, 10), pady=2, sticky="w")
                academic_section["Maths Mark"] = maths_entry

                cutoff_label = tk.Label(frame, text="Cut Off Mark:")
                cutoff_label.grid(row=len(labels) + 4, column=0, padx=(0, 5), sticky="w")
                cutoff_entry = tk.Entry(frame)
                cutoff_entry.grid(row=len(labels) + 4, column=1, padx=(0, 10), pady=2, sticky="w")
                academic_section["Cut Off Mark"] = cutoff_entry

                emis_label = tk.Label(frame, text="EMIS NO:")
                emis_label.grid(row=len(labels) + 5, column=0, padx=(0, 5), sticky="w")
                emis_entry = tk.Entry(frame)
                emis_entry.grid(row=len(labels) + 5, column=1, padx=(0, 10), pady=2, sticky="w")
                academic_section["EMIS NO"] = emis_entry

            self.academic_data.append(academic_section)
    def create_scholarship_details(self, window):
        # Set the background image for the window
        background_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your background image path
        background_label = tk.Label(window, image=background_image)
        background_label.photo = background_image  # Keep a reference to the PhotoImage
        background_label.place(relwidth=1, relheight=1)

        labels = ["PUDHUMAIPEN:", "PMMS:", "BC/MBC:", "NSP:", "7.5:", "FG:","CKCET:","PRIVATE:"]
        self.scholarship_data = {}

        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.place(x=20, y=20 + labels.index(label_text) * 30)  # Adjust the x and y coordinates as needed
            self.scholarship_data[label_text] = var

    def create_certificate_details(self, window):
        # Set the background image for the window
        background_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your background image path
        background_label = tk.Label(window, image=background_image)
        background_label.photo = background_image  # Keep a reference to the PhotoImage
        background_label.place(relwidth=1, relheight=1)

        labels = ["10th Original:", "11th Original:", "12th Original:", "TC:","Community (CC):", "Income (IC):", "FG:", "JD:", 
                   "Nativity (NT):","Father-Aadhar:", "Mother-Aadhar:", "Student-Aadhar:","Family/Smart card:", "Voter ID"]
        self.certificate_data = {}

        for label_text in labels:
            var = tk.StringVar(value="No")
            checkbox = tk.Checkbutton(window, text=label_text, variable=var, onvalue="Yes", offvalue="No")
            checkbox.place(x=20, y=20 + labels.index(label_text) * 30)  # Adjust the x and y coordinates as needed
            self.certificate_data[label_text] = var

    def create_bank_details(self, window):
        # Create a Label widget to display the background image
        # Load the background image
        bg_image = tk.PhotoImage(file="backgrounds/pic.ppm")  # Replace with your image file path

        # Create a Canvas widget to display the background image
        canvas = tk.Canvas(window, width=bg_image.width(), height=bg_image.height())
        canvas.pack()

        # Display the background image on the Canvas
        canvas.create_image(0, 0, anchor=tk.NW, image=bg_image)
        canvas.image = bg_image  # Keep a reference to the image

        labels = ["BANK NAME:","BRANCH NAME:", "ACCOUNT NUMBER:" ,"IFSC CODE:","MIRC CODE:", "FEES PAID", "FEES PENDING","DATE OF ADMISSION"]
        self.bank_data = {}

        for label_text in labels:
            label = tk.Label(window, text=label_text)
            label.place(x=20, y=20 + labels.index(label_text) * 30)  # Adjust the x and y coordinates as needed

            entry = tk.Entry(window)
            entry.place(x=200, y=20 + labels.index(label_text) * 30)  # Adjust the x coordinate as needed
            self.bank_data[label_text] = entry
            
        save_button = tk.Button(window, text="Save", command=self.save_data)
        save_button.place(x=200, y=280)
    def save_data(self):
        wb = openpyxl.load_workbook('main.xlsx')
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
        
        # Add a serial number (SNo) column, starting from 1
        serial_number = sheet.max_row + 1
        
        # Append the serial number to the row
        row.append(serial_number)
        
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

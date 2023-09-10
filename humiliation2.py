import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
class EditStudentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Edit Student Details")
        self.root.geometry('800x600')

        # Load student data from the actual saved data sheet
        self.student_data = self.load_student_data('student_details.xlsx')

        # Search section
        self.search_label = tk.Label(root, text="Search by:")
        self.search_label.pack()

        # Dropdown menu for search criteria
        self.search_criteria = tk.StringVar(root)
        self.search_criteria.set("Name")  # Default search criteria
        search_dropdown = ttk.Combobox(root, textvariable=self.search_criteria, values=["Name", "Application No"])
        search_dropdown.pack()

        # Entry for search input
        self.search_entry = tk.Entry(root)
        self.search_entry.pack()

        # Search button
        search_button = tk.Button(root, text="Search", command=self.search_student)
        search_button.pack()

        # Listbox to display matching student names
        self.student_listbox = tk.Listbox(root)
        self.student_listbox.pack()

        # Right-side panel for student photo
        self.photo_label = tk.Label(root, text="Student Photo")
        self.photo_label.pack(side="left")

        # Buttons for different student details (arranged horizontally)
        button_frame = tk.Frame(root)
        button_frame.pack(side="top")

        basic_details_button = tk.Button(button_frame, text="Student Basic Details", command=self.show_basic_details)
        family_details_button = tk.Button(button_frame, text="Family Details", command=self.show_family_details)
        academic_details_button = tk.Button(button_frame, text="Academic Details", command=self.show_academic_details)
        scholarship_details_button = tk.Button(button_frame, text="Scholarship Details", command=self.show_scholarship_details)
        bank_details_button = tk.Button(button_frame, text="Bank Details", command=self.show_bank_details)

        basic_details_button.pack(side="left", padx=5)
        family_details_button.pack(side="left", padx=5)
        academic_details_button.pack(side="left", padx=5)
        scholarship_details_button.pack(side="left", padx=5)
        bank_details_button.pack(side="left", padx=5)

        # Populate the student list with loaded data
        self.populate_student_list()

    def create_edit_window(self):
        # Create a new window for editing student details
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Student Details")
        edit_window.geometry('800x600')

        # Define a function to display student details and allow for updates
        def display_student_details(section_name, window):
            # Fetch the selected student's details from student_data dictionary
            selected_student = self.student_listbox.get(tk.ACTIVE)
            if selected_student:
                student_details = self.student_data[selected_student]

                # Create a new window for displaying and updating the section details
                section_window = tk.Toplevel(window)  # Pass the window as an argument
                section_window.title(section_name)

                # Create a dictionary to store the entry widgets for this section
                section_entries = {}

                # Function to update the section details
                def update_section_details():
                    # Update the student_details dictionary with new data from entry widgets
                    for field, entry in section_entries.items():
                        student_details[field] = entry.get()

                # Create entry widgets for each field in the section
                for field in student_details.get(section_name, {}):
                    label = tk.Label(section_window, text=field)
                    label.pack()
                    entry = tk.Entry(section_window)
                    entry.insert(0, student_details[section_name].get(field, ""))  # Pre-fill with existing data
                    entry.pack()
                    section_entries[field] = entry

                # Create an "Update" button for the section
                update_button = tk.Button(section_window, text="Update", command=update_section_details)
                update_button.pack()

        # Create buttons for different sections
        student_details_button = tk.Button(edit_window, text="Student Basic Details", command=lambda: display_student_details("Student Basic Details", edit_window))
        student_details_button.pack()

        family_details_button = tk.Button(edit_window, text="Family Details", command=lambda: display_student_details("Family Details", edit_window))
        family_details_button.pack()

        academic_details_button = tk.Button(edit_window, text="Academic Details", command=lambda: display_student_details("Academic Details", edit_window))
        academic_details_button.pack()

        scholarship_details_button = tk.Button(edit_window, text="Scholarship Details", command=lambda: display_student_details("Scholarship Details", edit_window))
        scholarship_details_button.pack()

        bank_details_button = tk.Button(edit_window, text="Bank Details", command=lambda: display_student_details("Bank Details", edit_window))
        bank_details_button.pack()

        certificate_details_button = tk.Button(edit_window, text="Certificate Details", command=lambda: display_student_details("Certificate Details", edit_window))
        certificate_details_button.pack()
    def load_student_data(self, filename):
        student_data = {}
        try:
            wb = openpyxl.load_workbook(filename)
            sheet = wb.active

            for row in sheet.iter_rows(min_row=2, values_only=True):
                s_no, application_no, date_of_application, name_of_student, aadhar_number, date_of_birth, gender, department, quota, community, caste, religion, mother_tongue, blood_group, medium_of_school, marital_status, father_name, mother_name, guardian_name, annual_income, parent_occupation, school_type, present_address, permanent_address, pincode, district, location_of_residence, parent_cell_number, student_cell_number, hostel, x_school_name, x_location, x_board, x_year_of_passing, x_percentage, xi_school_name, xi_location, xi_board, xi_year_of_passing, xi_percentage, xii_school_name, xii_location, xii_school_type, xii_board, xii_year_of_passing, xii_percentage, physics, chemistry, maths, cut_off_mark, emis_number, x_mark_sheet, xi_mark_sheet, xii_mark_sheet, tc, cc, ic, fg, jd, nt, aadhar_card_to_father, aadhar_card_to_mother, aadhar_card_to_student, family_card, voter_id,puthumaipen, pmms, bc_mbc, nsp, seven_point_five, fg_ckcet, private_bank_name, branch, account_number, ifsc_code, micr_code, fees_paid, fees_pending, date_of_admission = row

                student_data[name_of_student] = {
                    "S.No": s_no,
                    "Application No": application_no,
                    "Date of Application": date_of_application,
                    "Name of the Student": name_of_student,
                    "Aadhar Number": aadhar_number,
                    "Date of Birth": date_of_birth,
                    "Gender": gender,
                    "Department": department,
                    "Quota": quota,
                    "Community": community,
                    "Caste": caste,
                    "Religion": religion,
                    "Mother Tongue": mother_tongue,
                    "Blood Group": blood_group,
                    "Medium of School": medium_of_school,
                    "Marital Status": marital_status,
                    "Father Name": father_name,
                    "Mother Name": mother_name,
                    "Guardian Name": guardian_name,
                    "Annual Income": annual_income,
                    "Parent Occupation": parent_occupation,
                    "School Type": school_type,
                    "Present Address": present_address,
                    "Permanent Address": permanent_address,
                    "Pincode": pincode,
                    "District": district,
                    "Location of Residence": location_of_residence,
                    "Parent Cell Number": parent_cell_number,
                    "Student Cell Number": student_cell_number,
                    "Hostel": hostel,
                    "10th School Name": x_school_name,
                    "10th School Location": x_location,
                    "10th School Board": x_board,
                    "10th Year of Passing": x_year_of_passing,
                    "10th Percentage of Marks": x_percentage,
                    "11th School Name": xi_school_name,
                    "11th School Location": xi_location,
                    "11th School Board": xi_board,
                    "11th Year of Passing": xi_year_of_passing,
                    "11th Percentage of Marks": xi_percentage,
                    "12th School Name": xii_school_name,
                    "12th School Location": xii_location,
                    "12th School Type": xii_school_type,
                    "12th Board": xii_board,
                    "12th Year of Passing": xii_year_of_passing,
                    "12th Percentage of Marks": xii_percentage,
                    "Physics": physics,
                    "Chemistry": chemistry,
                    "Maths": maths,
                    "Cut Off Mark": cut_off_mark,
                    "EMIS Number": emis_number,
                    "X Mark Sheet": x_mark_sheet,
                    "XI Mark Sheet": xi_mark_sheet,
                    "XII Mark Sheet": xii_mark_sheet,
                    "TC": tc,
                    "CC": cc,
                    "IC": ic,
                    "FG": fg,
                    "JD": jd,
                    "NT": nt,
                    "Aadhar Card to Father": aadhar_card_to_father,
                    "Aadhar Card to Mother": aadhar_card_to_mother,
                    "Aadhar Card to Student": aadhar_card_to_student,
                    "Family Card": family_card,
                    "Voter ID": voter_id,
                    "Puthumaipen": puthumaipen,
                    "PMMS": pmms,
                    "BC/MBC": bc_mbc,
                    "NSP": nsp,
                    "7.5": seven_point_five,
                    "FG CKCET": fg_ckcet,
                    "Private Bank Name": private_bank_name,
                    "Branch": branch,
                    "Account Number": account_number,
                    "IFSC Code": ifsc_code,
                    "MICR Code": micr_code,
                    "Fees Paid": fees_paid,
                    "Fees Pending": fees_pending,
                    "Date of Admission": date_of_admission,
                }

            wb.close()
        except Exception as e:
            print(f"Error loading student data: {e}")

        return student_data


    def populate_student_list(self):
        # Populate the listbox with student names from the loaded data
        for student_name in self.student_data.keys():
            self.student_listbox.insert(tk.END, student_name)

    def search_student(self):
        # Implement search logic to populate the listbox with matching student names
        criteria = self.search_criteria.get()
        search_input = self.search_entry.get().strip()

        # Clear the listbox
        self.student_listbox.delete(0, tk.END)

        # Search for students based on criteria
        for name in self.student_data.keys():
            if criteria == "Name" and search_input.lower() in name.lower():
                self.student_listbox.insert(tk.END, name)
            elif criteria == "Application No" and search_input == self.student_data[name].get("Application No"):
                self.student_listbox.insert(tk.END, name)

    def show_basic_details(self):
        # Implement logic to display and edit basic student details
        selected_student = self.student_listbox.get(tk.ACTIVE)
        if selected_student:
            # Fetch and display basic details of the selected student
            # You can create a new window or frame to display and edit these details
            pass

    def show_family_details(self):
        # Implement logic to display and edit family details
        selected_student = self.student_listbox.get(tk.ACTIVE)
        if selected_student:
            # Fetch and display family details of the selected student
            # You can create a new window or frame to display and edit these details
            pass

    def show_academic_details(self):
        # Implement logic to display and edit academic details
        selected_student = self.student_listbox.get(tk.ACTIVE)
        if selected_student:
            # Fetch and display academic details of the selected student
            # You can create a new window or frame to display and edit these details
            pass

    def show_scholarship_details(self):
        # Implement logic to display and edit scholarship details
        selected_student = self.student_listbox.get(tk.ACTIVE)
        if selected_student:
            # Fetch and display scholarship details of the selected student
            # You can create a new window or frame to display and edit these details
            pass

    def show_bank_details(self):
        # Implement logic to display and edit bank details
        selected_student = self.student_listbox.get(tk.ACTIVE)
        if selected_student:
            # Fetch and display bank details of the selected student
            # You can create a new window or frame to display and edit these details
            pass


##if __name__ == "__main__":
##    root = tk.Tk()
##    app = EditStudentApp(root)
##    root.mainloop()

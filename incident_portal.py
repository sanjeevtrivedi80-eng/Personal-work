import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime

class IncidentPortal:
    def __init__(self, root):
        self.root = root
        self.root.title("ST Apps - IT Incident Portal")
        self.root.geometry("500x600")
        
        # Variables
        self.user_details = tk.StringVar()
        self.subject = tk.StringVar()
        self.attachment_path = tk.StringVar()
        self.db_file = "incident_records.xlsx"

        # Initialize Excel Database
        self.init_db()

        # --- UI Layout ---
        main_label = tk.Label(root, text="IT Incident Reporting Portal", font=("Arial", 16, "bold"), pady=20)
        main_label.pack()

        # User Details
        tk.Label(root, text="User Details (Name/ID):").pack(anchor="w", padx=50)
        tk.Entry(root, textvariable=self.user_details, width=50).pack(pady=5)

        # Subject
        tk.Label(root, text="Subject:").pack(anchor="w", padx=50)
        tk.Entry(root, textvariable=self.subject, width=50).pack(pady=5)

        # Description
        tk.Label(root, text="Description:").pack(anchor="w", padx=50)
        self.desc_text = tk.Text(root, height=8, width=37)
        self.desc_text.pack(pady=5)

        # Attachment Section
        attach_frame = tk.Frame(root)
        attach_frame.pack(pady=10)
        tk.Button(attach_frame, text="Attach File (JPG, PNG, PDF)", command=self.attach_file).grid(row=0, column=0)
        tk.Label(attach_frame, textvariable=self.attachment_path, fg="blue", wraplength=200).grid(row=0, column=1, padx=10)

        # Submit Button
        tk.Button(root, text="Submit Incident", bg="#2c3e50", fg="white", width=20, height=2, command=self.submit_incident).pack(pady=20)

    def init_db(self):
        """Creates the Excel file with headers if it doesn't exist."""
        if not os.path.exists(self.db_file):
            wb = Workbook()
            sheet = wb.active
            sheet.title = "Incidents"
            headers = ["Date", "User", "Subject", "Description", "Attachment Path"]
            sheet.append(headers)
            wb.save(self.db_file)

    def attach_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Image/PDF files", "*.jpg *.jpeg *.png *.pdf")]
        )
        if file_path:
            self.attachment_path.set(file_path)

    def submit_incident(self):
        # Validate inputs
        user = self.user_details.get()
        subj = self.subject.get()
        desc = self.desc_text.get("1.0", tk.END).strip()
        date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        attach = self.attachment_path.get()

        if not user or not subj or not desc:
            messagebox.showerror("Error", "Please fill in all required fields.")
            return

        try:
            # Append to Excel
            wb = openpyxl.load_workbook(self.db_file)
            sheet = wb.active
            sheet.append([date_str, user, subj, desc, attach])
            wb.save(self.db_file)

            messagebox.showinfo("Success", "Incident logged successfully!")
            
            # Clear fields
            self.user_details.set("")
            self.subject.set("")
            self.desc_text.delete("1.0", tk.END)
            self.attachment_path.set("")
            
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not save to Excel: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = IncidentPortal(root)
    root.mainloop()
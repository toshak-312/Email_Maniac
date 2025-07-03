import sys
import os
import csv
import smtplib
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Text, scrolledtext
from tkinter.font import Font
import re

class BulkEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Internship Request Email Sender")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Initialize variables
        self.csv_file_path = tk.StringVar()
        self.pdf_file_path = tk.StringVar()
        self.smtp_server = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.IntVar(value=587)
        self.email_account = tk.StringVar()
        self.email_password = tk.StringVar()
        self.email_subject = tk.StringVar(value="Internship Opportunity Request")
        self.csv_data = []
        self.column_headers = []
        self.selected_columns = {}
        self.preview_row = 0
        
        # Default template
        self.default_template = """Dear {{First Name}} {{Last Name}},

I hope this email finds you well. My name is [Your Name], and I am writing to express my interest in potential internship opportunities at {{Company Name}}, specifically in the {{Role}} department.

I am particularly interested in {{Company Name}} because of its innovative approach and reputation in the industry. I believe that my skills and enthusiasm would make me a valuable addition to your team.

I have attached my CV for your consideration. I would appreciate the opportunity to discuss how my background and skills would be a good fit for {{Company Name}}.

Thank you for considering my application. I look forward to the possibility of working with you.

Best regards,
[Your Name]
[Your Contact Information]
"""
        
        # Create UI
        self.create_ui()
        
    def create_ui(self):
        # Style configuration
        style = ttk.Style()
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TButton", background="#4285f4", foreground="black", font=('Arial', 10))
        style.configure("TLabel", background="#f5f5f5", font=('Arial', 10))
        style.map('TButton', background=[('active', '#5a95f5')])
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for tabbed interface
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Tab 1: Setup
        setup_tab = ttk.Frame(notebook)
        notebook.add(setup_tab, text="Setup")
        
        # Tab 2: Template
        template_tab = ttk.Frame(notebook)
        notebook.add(template_tab, text="Email Template")
        
        # Tab 3: Preview
        preview_tab = ttk.Frame(notebook)
        notebook.add(preview_tab, text="Preview")
        
        # Tab 4: Send
        send_tab = ttk.Frame(notebook)
        notebook.add(send_tab, text="Send Emails")
        
        # Setup Tab
        self.create_setup_tab(setup_tab)
        
        # Template Tab
        self.create_template_tab(template_tab)
        
        # Preview Tab
        self.create_preview_tab(preview_tab)
        
        # Send Tab
        self.create_send_tab(send_tab)
        
    def create_setup_tab(self, parent):
        # File Selection Frame
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # CSV File Selection
        ttk.Label(file_frame, text="CSV Data File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.csv_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_csv_file).grid(row=0, column=2, padx=5, pady=5)
        
        # CV/Resume File Selection
        ttk.Label(file_frame, text="CV/Resume File (PDF):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.pdf_file_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf_file).grid(row=1, column=2, padx=5, pady=5)
        
        # Email Settings Frame
        email_frame = ttk.LabelFrame(parent, text="Email Settings", padding="10")
        email_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # SMTP Settings
        ttk.Label(email_frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.smtp_server, width=30).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(email_frame, text="SMTP Port:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.smtp_port, width=30).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(email_frame, text="Email Account:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.email_account, width=30).grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(email_frame, text="App Password:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.email_password, width=30, show="*").grid(row=3, column=1, padx=5, pady=5)
        
        ttk.Label(email_frame, text="Email Subject:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.email_subject, width=50).grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        # CSV Mapping Frame (will be populated after CSV is loaded)
        self.mapping_frame = ttk.LabelFrame(parent, text="CSV Column Mapping", padding="10")
        self.mapping_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(self.mapping_frame, text="Please load a CSV file to map columns").pack(pady=20)
        
    def create_template_tab(self, parent):
        # Template Frame
        template_frame = ttk.Frame(parent, padding="10")
        template_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(template_frame, text="Email Template (Use {{Column Name}} for placeholders):").pack(anchor=tk.W, pady=(0, 5))
        
        # Template text area with scrollbar
        self.template_text = scrolledtext.ScrolledText(template_frame, wrap=tk.WORD, width=80, height=20, font=("Arial", 11))
        self.template_text.pack(fill=tk.BOTH, expand=True)
        self.template_text.insert(tk.END, self.default_template)
        
        # Template hints
        hint_text = "Template Hints:\n"
        hint_text += "- Use {{Column Name}} to insert values from your CSV.\n"
        hint_text += "- Example: 'Dear {{First Name}}' will be replaced with actual first names.\n"
        hint_text += "- Customize the template to suit your specific needs."
        
        hint_label = ttk.Label(template_frame, text=hint_text, justify=tk.LEFT, background="#e6f3ff", padding=10)
        hint_label.pack(fill=tk.X, pady=10)
        
        # Buttons frame
        button_frame = ttk.Frame(template_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Reset to Default", command=self.reset_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save Template", command=self.save_template).pack(side=tk.LEFT, padx=5)
    
    def create_preview_tab(self, parent):
        # Preview Frame
        preview_frame = ttk.Frame(parent, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Navigation Frame
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(nav_frame, text="Preview Email:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(nav_frame, text="Previous", command=self.prev_preview).pack(side=tk.LEFT, padx=5)
        
        self.current_preview_label = ttk.Label(nav_frame, text="No data loaded")
        self.current_preview_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(nav_frame, text="Next", command=self.next_preview).pack(side=tk.LEFT, padx=5)
        
        # Preview text area
        ttk.Label(preview_frame, text="Preview:").pack(anchor=tk.W, pady=(10, 5))
        self.preview_area = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD, width=80, height=25, font=("Arial", 11))
        self.preview_area.pack(fill=tk.BOTH, expand=True)
        self.preview_area.insert(tk.END, "Load CSV data and set up template to see preview")
        self.preview_area.config(state=tk.DISABLED)
        
    def create_send_tab(self, parent):
        # Send Frame
        send_frame = ttk.Frame(parent, padding="10")
        send_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Send options
        options_frame = ttk.Frame(send_frame)
        options_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(options_frame, text="Send Options:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.test_mode_var = tk.BooleanVar(value=True)
        test_check = ttk.Checkbutton(options_frame, text="Test Mode (Send to yourself only)", variable=self.test_mode_var)
        test_check.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        self.progress_frame = ttk.Frame(send_frame)
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(self.progress_frame, text="Progress:").pack(anchor=tk.W, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.progress_label = ttk.Label(self.progress_frame, text="Ready to send")
        self.progress_label.pack(anchor=tk.W, pady=5)
        
        # Log area
        ttk.Label(send_frame, text="Log:").pack(anchor=tk.W, pady=(10, 5))
        self.log_area = scrolledtext.ScrolledText(send_frame, wrap=tk.WORD, width=80, height=15, font=("Arial", 10))
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.config(state=tk.DISABLED)
        
        # Send button
        send_button_frame = ttk.Frame(send_frame)
        send_button_frame.pack(fill=tk.X, pady=10)
        
        self.send_button = ttk.Button(send_button_frame, text="Send Emails", command=self.send_emails)
        self.send_button.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(send_button_frame, text="Verify Settings", command=self.verify_settings).pack(side=tk.RIGHT, padx=5)
    
    def browse_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.csv_file_path.set(file_path)
            self.load_csv_data()
    
    def browse_pdf_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if file_path:
            self.pdf_file_path.set(file_path)
    
    def load_csv_data(self):
        try:
            # Clear existing widgets in mapping frame
            for widget in self.mapping_frame.winfo_children():
                widget.destroy()
            
            with open(self.csv_file_path.get(), 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                self.column_headers = next(reader)
                self.csv_data = list(reader)
            
            if not self.csv_data:
                messagebox.showwarning("Warning", "The CSV file appears to be empty!")
                return
            
            # Create dropdown mapping for each column
            ttk.Label(self.mapping_frame, text="Map CSV columns to template variables:").grid(
                row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
            
            # Create required field mappings
            required_fields = ["First Name", "Last Name", "Company Name", "Role", "Email"]
            self.selected_columns = {}
            
            for i, field in enumerate(required_fields):
                ttk.Label(self.mapping_frame, text=f"{field}:").grid(row=i+1, column=0, sticky=tk.W, pady=5)
                combo = ttk.Combobox(self.mapping_frame, values=self.column_headers, width=30)
                combo.grid(row=i+1, column=1, sticky=tk.W, padx=5, pady=5)
                
                # Try to auto-select matching columns
                for header in self.column_headers:
                    if field.lower() in header.lower():
                        combo.set(header)
                        break
                
                self.selected_columns[field] = combo
            
            # Show sample data
            ttk.Label(self.mapping_frame, text="Sample Data (First Row):").grid(
                row=len(required_fields)+1, column=0, columnspan=2, sticky=tk.W, pady=(15, 5))
            
            sample_text = ""
            for i, header in enumerate(self.column_headers):
                if i < len(self.csv_data[0]):
                    sample_text += f"{header}: {self.csv_data[0][i]}\n"
            
            sample_area = scrolledtext.ScrolledText(self.mapping_frame, wrap=tk.WORD, width=60, height=8)
            sample_area.grid(row=len(required_fields)+2, column=0, columnspan=2, sticky=tk.W, pady=5)
            sample_area.insert(tk.END, sample_text)
            sample_area.config(state=tk.DISABLED)
            
            # Update preview
            self.update_preview()
            
            # Log
            self.log(f"Loaded CSV file with {len(self.csv_data)} records and {len(self.column_headers)} columns")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
    
    def reset_template(self):
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset to the default template?"):
            self.template_text.delete(1.0, tk.END)
            self.template_text.insert(tk.END, self.default_template)
    
    def save_template(self):
        # Save template functionality could be expanded to save to file
        messagebox.showinfo("Template Saved", "Template has been saved.")
    
    def update_preview(self):
        if not self.csv_data:
            return
        
        # Update navigation label
        self.current_preview_label.config(text=f"Record {self.preview_row + 1} of {len(self.csv_data)}")
        
        # Get current template
        template = self.template_text.get(1.0, tk.END)
        
        # Replace placeholders with values from the current row
        email_body = self.generate_email_content(template, self.preview_row)
        
        # Update preview area
        self.preview_area.config(state=tk.NORMAL)
        self.preview_area.delete(1.0, tk.END)
        self.preview_area.insert(tk.END, f"Subject: {self.email_subject.get()}\n\n")
        self.preview_area.insert(tk.END, email_body)
        
        # Add recipient info
        email_col_index = None
        for i, header in enumerate(self.column_headers):
            for field, combo in self.selected_columns.items():
                if combo.get() == header and field == "Email":
                    email_col_index = i
                    break
        
        if email_col_index is not None and email_col_index < len(self.csv_data[self.preview_row]):
            recipient = self.csv_data[self.preview_row][email_col_index]
            self.preview_area.insert(tk.END, f"\n\n[Will be sent to: {recipient}]")
        
        self.preview_area.config(state=tk.DISABLED)
    
    def prev_preview(self):
        if self.csv_data:
            self.preview_row = (self.preview_row - 1) % len(self.csv_data)
            self.update_preview()
    
    def next_preview(self):
        if self.csv_data:
            self.preview_row = (self.preview_row + 1) % len(self.csv_data)
            self.update_preview()
    
    def generate_email_content(self, template, row_index):
        if not self.csv_data or row_index >= len(self.csv_data):
            return "No data available"
        
        row = self.csv_data[row_index]
        email_content = template
        
        # For each column mapping, replace placeholders
        for field, combo in self.selected_columns.items():
            selected_header = combo.get()
            if selected_header:
                try:
                    col_index = self.column_headers.index(selected_header)
                    value = row[col_index] if col_index < len(row) else ""
                    email_content = email_content.replace(f"{{{{{field}}}}}", value)
                except ValueError:
                    pass
        
        # Also look for other column headers that might be in the template
        for i, header in enumerate(self.column_headers):
            placeholder = f"{{{{{header}}}}}"
            if placeholder in email_content and i < len(row):
                email_content = email_content.replace(placeholder, row[i])
        
        return email_content
    
    def verify_settings(self):
        missing = []
        
        # Check files
        if not self.csv_file_path.get():
            missing.append("CSV file")
        if not self.pdf_file_path.get():
            missing.append("PDF Resume/CV file")
        
        # Check email settings
        if not self.email_account.get():
            missing.append("Email account")
        if not self.email_password.get():
            missing.append("Email password")
        
        # Check mappings if CSV is loaded
        if self.csv_data:
            for field, combo in self.selected_columns.items():
                if not combo.get():
                    missing.append(f"{field} column mapping")
        
        if missing:
            messagebox.showwarning("Missing Settings", 
                                  f"The following settings are missing:\n- " + "\n- ".join(missing))
        else:
            messagebox.showinfo("Settings Verified", "All required settings are provided!")
    
    def send_emails(self):
        # Verify settings first
        if not self.csv_data:
            messagebox.showerror("Error", "No CSV data loaded")
            return
        
        # Check required fields
        missing = []
        for field, combo in self.selected_columns.items():
            if not combo.get():
                missing.append(field)
        
        if missing:
            messagebox.showerror("Error", f"Missing mappings for: {', '.join(missing)}")
            return
        
        if not self.pdf_file_path.get() or not os.path.exists(self.pdf_file_path.get()):
            messagebox.showerror("Error", "Invalid PDF file path")
            return
        
        if not self.email_account.get() or not self.email_password.get():
            messagebox.showerror("Error", "Email credentials are required")
            return
        
        # Get email column index
        email_col_index = None
        for i, header in enumerate(self.column_headers):
            if header == self.selected_columns["Email"].get():
                email_col_index = i
                break
        
        if email_col_index is None:
            messagebox.showerror("Error", "Email column not properly mapped")
            return
        
        # Prepare for sending
        self.send_button.config(state=tk.DISABLED)
        self.progress_bar["maximum"] = len(self.csv_data)
        self.progress_bar["value"] = 0
        self.progress_label.config(text="Preparing to send emails...")
        
        # Clear log
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        # Get template
        template = self.template_text.get(1.0, tk.END)
        
        # Check if test mode is enabled
        test_mode = self.test_mode_var.get()
        if test_mode:
            self.log("TEST MODE ENABLED - Emails will be sent to your own address only")
        
        # Start sending in a separate thread
        threading.Thread(target=self.send_emails_thread, 
                        args=(template, email_col_index, test_mode)).start()
    
    def send_emails_thread(self, template, email_col_index, test_mode):
        try:
            # Connect to SMTP server
            self.log(f"Connecting to SMTP server {self.smtp_server.get()}:{self.smtp_port.get()}...")
            
            server = smtplib.SMTP(self.smtp_server.get(), self.smtp_port.get())
            server.starttls()
            server.login(self.email_account.get(), self.email_password.get())
            
            self.log("Connected to SMTP server successfully")
            
            # Send emails
            success_count = 0
            failed_count = 0
            
            for i, row in enumerate(self.csv_data):
                try:
                    # Update progress
                    self.progress_bar["value"] = i + 1
                    self.progress_label.config(text=f"Sending email {i+1} of {len(self.csv_data)}...")
                    
                    # Get recipient email
                    if email_col_index < len(row):
                        recipient_email = row[email_col_index].strip()
                    else:
                        self.log(f"Skipping row {i+1}: Missing email address")
                        failed_count += 1
                        continue
                    
                    # Validate email
                    if not self.is_valid_email(recipient_email):
                        self.log(f"Skipping row {i+1}: Invalid email address: {recipient_email}")
                        failed_count += 1
                        continue
                    
                    # Generate email content
                    email_body = self.generate_email_content(template, i)
                    
                    # Create message
                    msg = MIMEMultipart()
                    msg['From'] = self.email_account.get()
                    
                    # If test mode, send to own email
                    if test_mode:
                        msg['To'] = self.email_account.get()
                        test_info = f"\n\n[TEST MODE - Original recipient would have been: {recipient_email}]"
                        email_body += test_info
                    else:
                        msg['To'] = recipient_email
                    
                    msg['Subject'] = self.email_subject.get()
                    
                    # Attach email body
                    msg.attach(MIMEText(email_body, 'plain'))
                    
                    # Attach CV/Resume
                    with open(self.pdf_file_path.get(), 'rb') as f:
                        attachment = MIMEApplication(f.read(), _subtype="pdf")
                    
                    attachment.add_header('Content-Disposition', 'attachment', 
                                        filename=os.path.basename(self.pdf_file_path.get()))
                    msg.attach(attachment)
                    
                    # Send email
                    server.send_message(msg)
                    
                    # Log success
                    if test_mode:
                        self.log(f"Test email sent successfully for row {i+1} ({recipient_email})")
                    else:
                        self.log(f"Email sent successfully to: {recipient_email}")
                    
                    success_count += 1
                    
                    # Slight delay to avoid rate limiting
                    self.root.after(100)
                    
                except Exception as e:
                    self.log(f"Failed to send email for row {i+1}: {str(e)}")
                    failed_count += 1
            
            # Close connection
            server.quit()
            self.log(f"Email sending completed. Success: {success_count}, Failed: {failed_count}")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.send_button.config(state=tk.NORMAL)
            self.progress_label.config(text="Email sending process completed")
    
    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = BulkEmailSender(root)
    root.mainloop()

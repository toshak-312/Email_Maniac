import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
import json
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import threading
import time

class InternshipEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Internship Email Sender")
        self.root.geometry("900x700")
        
        # Default settings
        self.default_settings = {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "email_subject": "Internship Application",
            "column_mappings": {
                "first_name": "",
                "last_name": "",
                "company": "",
                "role": "",
                "email": ""
            },
            "email_template": "Dear {first_name},\n\nI hope this email finds you well. My name is [YOUR NAME], and I am writing to express my interest in the {role} position at {company}.\n\nI am currently pursuing [YOUR DEGREE] and am eager to apply my skills in a real-world setting. I have attached my CV for your consideration.\n\nThank you for your time and consideration. I look forward to hearing from you.\n\nBest regards,\n[YOUR NAME]"
        }
        
        # App state
        self.settings = self.load_settings()
        self.cv_file_path = self.settings.get("cv_file_path", "")
        self.cv_file_data = self.settings.get("cv_file_data", "")
        self.cv_file_name = self.settings.get("cv_file_name", "")
        self.csv_file_path = ""
        self.df = None
        self.selected_columns = {}
        self.email_pattern = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Setup tabs
        self.setup_csv_tab()
        self.setup_template_tab()
        self.setup_email_tab()
        self.setup_send_tab()
        
        # Load saved CV if it exists
        if self.cv_file_data and self.cv_file_name:
            self.cv_label.config(text=f"CV loaded: {self.cv_file_name}")
    
    def load_settings(self):
        """Load settings from file if exists, otherwise use defaults"""
        try:
            if os.path.exists("email_sender_settings.json"):
                with open("email_sender_settings.json", "r") as file:
                    settings = json.load(file)
                return settings
            return self.default_settings.copy()
        except Exception as e:
            print(f"Error loading settings: {e}")
            return self.default_settings.copy()
    
    def save_settings(self):
        """Save current settings to file"""
        try:
            # Prepare settings to save
            settings_to_save = {
                "smtp_server": self.smtp_server_entry.get(),
                "smtp_port": self.smtp_port_entry.get(),
                "email_subject": self.subject_entry.get(),
                "column_mappings": {k: v.get() for k, v in self.column_entries.items()},
                "email_template": self.template_text.get("1.0", tk.END),
                "cv_file_path": self.cv_file_path,
                "cv_file_data": self.cv_file_data,
                "cv_file_name": self.cv_file_name
            }
            
            with open("email_sender_settings.json", "w") as file:
                json.dump(settings_to_save, file)
                
            messagebox.showinfo("Settings Saved", "Your settings have been saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")
    
    def setup_csv_tab(self):
        """Setup the CSV tab for file selection and column mapping"""
        csv_tab = ttk.Frame(self.notebook)
        self.notebook.add(csv_tab, text="1. CSV Data")
        
        # CSV File Selection
        frame = ttk.LabelFrame(csv_tab, text="CSV File")
        frame.pack(fill=tk.X, expand=False, padx=10, pady=10)
        
        ttk.Button(frame, text="Select CSV File", command=self.select_csv_file).pack(side=tk.LEFT, padx=5, pady=5)
        self.csv_label = ttk.Label(frame, text="No file selected")
        self.csv_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Column Mapping
        mapping_frame = ttk.LabelFrame(csv_tab, text="Column Mapping")
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create column mapping entries
        self.column_entries = {}
        row = 0
        
        # First Name (Default and required)
        ttk.Label(mapping_frame, text="First Name Column:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        self.column_entries["first_name"] = tk.StringVar(value=self.settings["column_mappings"]["first_name"])
        ttk.Entry(mapping_frame, textvariable=self.column_entries["first_name"]).grid(row=row, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        row += 1
        
        # Optional mappings with add/remove functionality
        self.optional_mappings = ["last_name", "company", "role", "email"]
        self.optional_frames = {}
        
        for field in self.optional_mappings:
            self.optional_frames[field] = ttk.Frame(mapping_frame)
            self.optional_frames[field].grid(row=row, column=0, columnspan=3, sticky=tk.W+tk.E, padx=5, pady=2)
            
            ttk.Label(self.optional_frames[field], text=f"{field.replace('_', ' ').title()} Column:").pack(side=tk.LEFT, padx=5)
            self.column_entries[field] = tk.StringVar(value=self.settings["column_mappings"][field])
            ttk.Entry(self.optional_frames[field], textvariable=self.column_entries[field]).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            
            # If there's a value saved for this field, show it
            if self.settings["column_mappings"][field]:
                self.optional_frames[field].grid()
            else:
                self.optional_frames[field].grid_remove()
            
            row += 1
        
        # Add/Remove Field buttons
        button_frame = ttk.Frame(mapping_frame)
        button_frame.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=10)
        
        ttk.Button(button_frame, text="+ Add Field", command=self.add_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="- Remove Field", command=self.remove_field).pack(side=tk.LEFT, padx=5)
        
        # Data preview
        preview_frame = ttk.LabelFrame(csv_tab, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.preview_tree = ttk.Treeview(preview_frame)
        self.preview_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_tree.configure(yscrollcommand=scrollbar.set)
        
        # Next button
        ttk.Button(csv_tab, text="Next >", command=lambda: self.notebook.select(1)).pack(side=tk.RIGHT, padx=10, pady=10)
    
    def add_field(self):
        """Add an optional field to the column mapping"""
        # Find the first hidden field and show it
        for field in self.optional_mappings:
            if not self.optional_frames[field].winfo_viewable():
                self.optional_frames[field].grid()
                break
    
    def remove_field(self):
        """Remove the last visible optional field"""
        # Find the last visible field and hide it
        visible_fields = [f for f in self.optional_mappings if self.optional_frames[f].winfo_viewable()]
        if visible_fields:
            field = visible_fields[-1]
            self.column_entries[field].set("")  # Clear the value
            self.optional_frames[field].grid_remove()
    
    def setup_template_tab(self):
        """Setup the template tab for email content"""
        template_tab = ttk.Frame(self.notebook)
        self.notebook.add(template_tab, text="2. Email Template")
        
        # Template editor
        frame = ttk.LabelFrame(template_tab, text="Email Template")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Help text
        help_text = "Use placeholders like {first_name}, {last_name}, {company}, {role}, {email} in your template."
        ttk.Label(frame, text=help_text).pack(anchor=tk.W, padx=5, pady=5)
        
        # Template text area
        self.template_text = scrolledtext.ScrolledText(frame, height=20)
        self.template_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.template_text.insert(tk.END, self.settings["email_template"])
        
        # CV attachment
        cv_frame = ttk.LabelFrame(template_tab, text="CV Attachment")
        cv_frame.pack(fill=tk.X, expand=False, padx=10, pady=10)
        
        ttk.Button(cv_frame, text="Select CV (PDF)", command=self.select_cv).pack(side=tk.LEFT, padx=5, pady=5)
        self.cv_label = ttk.Label(cv_frame, text="No CV selected")
        self.cv_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Navigation buttons
        button_frame = ttk.Frame(template_tab)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=10)
        
        ttk.Button(button_frame, text="< Back", command=lambda: self.notebook.select(0)).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Next >", command=lambda: self.notebook.select(2)).pack(side=tk.RIGHT)
    
    def setup_email_tab(self):
        """Setup the email tab for email settings"""
        email_tab = ttk.Frame(self.notebook)
        self.notebook.add(email_tab, text="3. Email Settings")
        
        # Email settings
        settings_frame = ttk.LabelFrame(email_tab, text="Email Settings")
        settings_frame.pack(fill=tk.X, expand=False, padx=10, pady=10)
        
        # From email
        ttk.Label(settings_frame, text="Your Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.from_email_entry = ttk.Entry(settings_frame, width=40)
        self.from_email_entry.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        # App password
        ttk.Label(settings_frame, text="App Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.password_entry = ttk.Entry(settings_frame, width=40, show="*")
        self.password_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        # SMTP settings
        ttk.Label(settings_frame, text="SMTP Server:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.smtp_server_entry = ttk.Entry(settings_frame, width=40)
        self.smtp_server_entry.grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.smtp_server_entry.insert(0, self.settings["smtp_server"])
        
        ttk.Label(settings_frame, text="SMTP Port:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.smtp_port_entry = ttk.Entry(settings_frame, width=10)
        self.smtp_port_entry.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        self.smtp_port_entry.insert(0, self.settings["smtp_port"])
        
        # Email subject
        ttk.Label(settings_frame, text="Email Subject:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.subject_entry = ttk.Entry(settings_frame, width=40)
        self.subject_entry.grid(row=4, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        self.subject_entry.insert(0, self.settings["email_subject"])
        
        # Save settings button
        ttk.Button(settings_frame, text="Save Settings", command=self.save_settings).grid(row=5, column=0, columnspan=2, pady=10)
        
        # Navigation buttons
        button_frame = ttk.Frame(email_tab)
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=10)
        
        ttk.Button(button_frame, text="< Back", command=lambda: self.notebook.select(1)).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Next >", command=lambda: self.notebook.select(3)).pack(side=tk.RIGHT)
    
    def setup_send_tab(self):
        """Setup the send tab for reviewing and sending emails"""
        send_tab = ttk.Frame(self.notebook)
        self.notebook.add(send_tab, text="4. Send Emails")
        
        # Review section
        review_frame = ttk.LabelFrame(send_tab, text="Review")
        review_frame.pack(fill=tk.X, expand=False, padx=10, pady=10)
        
        # Summary of what will be sent
        self.summary_text = scrolledtext.ScrolledText(review_frame, height=8, width=80)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Send button
        send_button_frame = ttk.Frame(send_tab)
        send_button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.send_button = ttk.Button(send_button_frame, text="Send Emails", command=self.send_emails)
        self.send_button.pack(side=tk.RIGHT)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(send_tab, text="Progress")
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.progress_text = scrolledtext.ScrolledText(progress_frame, height=15)
        self.progress_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Navigation button
        ttk.Button(send_tab, text="< Back", command=lambda: self.notebook.select(2)).pack(side=tk.LEFT, padx=10, pady=10)
    
    def select_csv_file(self):
        """Handle CSV file selection"""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.csv_file_path = file_path
                self.csv_label.config(text=os.path.basename(file_path))
                
                # Load and preview the data
                self.df = pd.read_csv(file_path)
                self.update_preview()
                
                # Auto-detect column mappings
                self.auto_detect_columns()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV file: {e}")
    
    def update_preview(self):
        """Update the preview treeview with CSV data"""
        if self.df is None:
            return
        
        # Clear existing data
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        # Configure columns
        self.preview_tree["columns"] = list(self.df.columns)
        self.preview_tree["show"] = "headings"
        
        for col in self.df.columns:
            self.preview_tree.heading(col, text=col)
            # Adjust column width based on data
            max_width = max(len(str(col)), self.df[col].astype(str).str.len().max())
            self.preview_tree.column(col, width=min(max_width * 10, 200))
        
        # Add data rows (limiting to first 10 for performance)
        for idx, row in self.df.head(10).iterrows():
            self.preview_tree.insert("", "end", values=list(row))
    
    def auto_detect_columns(self):
        """Auto-detect column mappings based on common naming patterns"""
        if self.df is None:
            return
        
        columns = self.df.columns
        mappings = {
            "first_name": ["first name", "firstname", "first", "fname", "given name", "name"],
            "last_name": ["last name", "lastname", "last", "lname", "surname", "family name"],
            "company": ["company", "company name", "employer", "organization", "organisation", "firm"],
            "role": ["role", "position", "job title", "title", "job", "designation"],
            "email": ["email", "email address", "e-mail", "mail"]
        }
        
        # Look for exact or partial matches
        for field, patterns in mappings.items():
            for col in columns:
                col_lower = col.lower()
                # Check for exact match first
                if col_lower in patterns:
                    self.column_entries[field].set(col)
                    break
                # Then check for partial matches
                for pattern in patterns:
                    if pattern in col_lower:
                        self.column_entries[field].set(col)
                        break
        
        # Special case for email - try to auto-detect using regex pattern
        if not self.column_entries["email"].get() and self.df is not None:
            for col in columns:
                # Check if column contains values that look like email addresses
                sample = self.df[col].astype(str).iloc[:5].tolist()  # Check first 5 rows
                if any(self.email_pattern.match(str(val)) for val in sample):
                    self.column_entries["email"].set(col)
                    break
        
        # Make sure all the visible optional fields have their frames shown
        for field in self.optional_mappings:
            if self.column_entries[field].get():
                self.optional_frames[field].grid()
    
    def select_cv(self):
        """Handle CV file selection"""
        file_path = filedialog.askopenfilename(
            title="Select CV (PDF)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Read file as binary and encode to base64 for storage
                with open(file_path, "rb") as file:
                    file_data = file.read()
                    self.cv_file_data = base64.b64encode(file_data).decode("utf-8")
                
                self.cv_file_path = file_path
                self.cv_file_name = os.path.basename(file_path)
                self.cv_label.config(text=f"CV loaded: {self.cv_file_name}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CV file: {e}")
    
    def update_summary(self):
        """Update the summary text before sending emails"""
        if self.df is None:
            self.summary_text.delete("1.0", tk.END)
            self.summary_text.insert(tk.END, "Please load a CSV file first.")
            return False
        
        # Get column mappings
        column_mappings = {k: v.get() for k, v in self.column_entries.items() if v.get()}
        
        # Validate required columns
        if not column_mappings.get("first_name"):
            self.summary_text.delete("1.0", tk.END)
            self.summary_text.insert(tk.END, "First Name column is required. Please set it in the CSV Data tab.")
            return False
        
        # Get email column - it's required
        email_col = column_mappings.get("email")
        if not email_col:
            self.summary_text.delete("1.0", tk.END)
            self.summary_text.insert(tk.END, "Email column is required. Please set it in the CSV Data tab.")
            return False
        
        # Count valid emails
        valid_emails = 0
        for email in self.df[email_col].astype(str):
            if self.email_pattern.match(email):
                valid_emails += 1
        
        # Build summary
        summary = f"""Ready to send {valid_emails} personalized emails.

From: {self.from_email_entry.get()}
Subject: {self.subject_entry.get()}
CV Attachment: {"Yes" if self.cv_file_data else "No"}

Recipients will be read from: {os.path.basename(self.csv_file_path)}
Using columns: {', '.join([f'{k} ({v})' for k, v in column_mappings.items()])}

Press 'Send Emails' to start sending."""
        
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert(tk.END, summary)
        return True
    
    def send_emails(self):
        """Start the email sending process"""
        # Update and validate summary first
        if not self.update_summary():
            return
        
        # Validate email settings
        if not self.from_email_entry.get():
            messagebox.showerror("Error", "Please enter your email address.")
            return
        
        if not self.password_entry.get():
            messagebox.showerror("Error", "Please enter your app password.")
            return
        
        # Get email column
        email_col = self.column_entries["email"].get()
        if not email_col:
            messagebox.showerror("Error", "Email column is required.")
            return
        
        # Start sending in a separate thread
        self.send_button.config(state="disabled")
        self.progress_text.delete("1.0", tk.END)
        self.progress_bar["value"] = 0
        
        threading.Thread(target=self.send_emails_task, daemon=True).start()
    
    def send_emails_task(self):
        """Task to send emails in background thread"""
        try:
            # Get column mappings
            column_mappings = {k: v.get() for k, v in self.column_entries.items() if v.get()}
            
            # Prepare for sending
            email_col = column_mappings.get("email")
            template = self.template_text.get("1.0", tk.END)
            subject = self.subject_entry.get()
            from_email = self.from_email_entry.get()
            password = self.password_entry.get()
            smtp_server = self.smtp_server_entry.get()
            smtp_port = int(self.smtp_port_entry.get())
            
            # Get list of recipients with valid emails
            recipients = []
            for _, row in self.df.iterrows():
                email = str(row[email_col])
                if self.email_pattern.match(email):
                    # Create dict of placeholders for this recipient
                    placeholders = {}
                    for field, col in column_mappings.items():
                        if col:  # If column is mapped
                            placeholders[field] = str(row[col])
                    
                    recipients.append({
                        "email": email,
                        "placeholders": placeholders
                    })
            
            if not recipients:
                self.log_progress("No valid email addresses found in the CSV.")
                self.send_button.config(state="normal")
                return
            
            # Set progress bar maximum
            total = len(recipients)
            self.progress_bar["maximum"] = total
            
            # Connect to SMTP server
            self.log_progress(f"Connecting to SMTP server {smtp_server}:{smtp_port}...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(from_email, password)
            self.log_progress("Connected to SMTP server.")
            
            # Send emails
            success_count = 0
            for i, recipient in enumerate(recipients):
                try:
                    # Personalize the email
                    personalized_message = template
                    for placeholder, value in recipient["placeholders"].items():
                        personalized_message = personalized_message.replace(f"{{{placeholder}}}", value)
                    
                    # Create message
                    msg = MIMEMultipart()
                    msg["From"] = from_email
                    msg["To"] = recipient["email"]
                    msg["Subject"] = subject
                    
                    # Add body
                    msg.attach(MIMEText(personalized_message, "plain"))
                    
                    # Add CV if available
                    if self.cv_file_data:
                        attachment = MIMEApplication(base64.b64decode(self.cv_file_data))
                        attachment.add_header("Content-Disposition", "attachment", filename=self.cv_file_name)
                        msg.attach(attachment)
                    
                    # Send the email
                    server.send_message(msg)
                    
                    # Log progress
                    name = recipient["placeholders"].get("first_name", "")
                    log_msg = f"Sent to {name} ({recipient['email']})"
                    self.log_progress(log_msg)
                    success_count += 1
                    
                except Exception as e:
                    self.log_progress(f"Failed to send to {recipient['email']}: {e}")
                
                # Update progress bar
                self.progress_bar["value"] = i + 1
                self.root.update_idletasks()
                
                # Sleep briefly to avoid rate limiting
                time.sleep(0.1)
            
            # Close connection
            server.quit()
            
            # Final summary
            self.log_progress(f"\nFinished sending emails. Success: {success_count}/{total}")
            
        except Exception as e:
            self.log_progress(f"Error: {e}")
        
        finally:
            # Re-enable send button
            self.send_button.config(state="normal")
    
    def log_progress(self, message):
        """Log progress message to progress text widget"""
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = InternshipEmailSender(root)
    root.mainloop()

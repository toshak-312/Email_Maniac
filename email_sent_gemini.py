import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os
import re # For email validation

CONFIG_FILE = "bulk_emailer_config.json"
DEFAULT_PLACEHOLDERS = {
    "FIRST_NAME": "{{FIRST_NAME}}",
    "LAST_NAME": "{{LAST_NAME}}",
    "COMPANY_NAME": "{{COMPANY_NAME}}",
    "ROLE": "{{ROLE}}",
}

# Heuristics for auto-detecting columns
AUTO_DETECT_PATTERNS = {
    "email_column": ['email', 'e-mail', 'email address', 'e mail'],
    "FIRST_NAME": ['firstname', 'first name', 'given name', 'first'],
    "COMPANY_NAME": ['companyname', 'company name', 'organization', 'company', 'employer'],
}


class BulkEmailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Internship Emailer V2")
        self.root.geometry("900x700") # Adjusted for tabs

        # --- Data ---
        self.csv_file_path = tk.StringVar()
        self.cv_file_path = tk.StringVar()
        self.csv_headers = []
        self.csv_data = []

        # --- Email Column ---
        self.email_column_var = tk.StringVar()

        # --- Column Mappings ---
        self.column_mappings = {key: tk.StringVar() for key in DEFAULT_PLACEHOLDERS}

        # --- Email Content ---
        self.email_subject_var = tk.StringVar()
        self.email_body_text_widget = None # Will be tk.Text widget

        # --- SMTP Settings ---
        self.smtp_email_var = tk.StringVar()
        self.smtp_password_var = tk.StringVar()

        self.load_config() # Load config before creating widgets that depend on it
        self.create_widgets()
        self.update_column_mapping_dropdowns_state()

        # If CSV path was loaded from config, load its data and try auto-map
        if self.csv_file_path.get() and os.path.exists(self.csv_file_path.get()):
            self._load_csv_data(self.csv_file_path.get(), silent=False)


    def save_config(self):
        """Saves current settings to the config file."""
        if not hasattr(self, 'email_body_text_widget') or self.email_body_text_widget is None:
             # This can happen if save_config is called before UI fully initialized (e.g. on_closing early)
             # For now, we'll try to get it from default_email_body if widget not ready
             email_body_content = getattr(self, 'default_email_body', "")
        else:
            email_body_content = self.email_body_text_widget.get("1.0", tk.END).strip()

        config_data = {
            "csv_file_path": self.csv_file_path.get(),
            "cv_file_path": self.cv_file_path.get(),
            "email_column": self.email_column_var.get(),
            "column_mappings": {key: var.get() for key, var in self.column_mappings.items()},
            "email_subject": self.email_subject_var.get(),
            "email_body": email_body_content,
            "smtp_email": self.smtp_email_var.get(),
            "smtp_password": self.smtp_password_var.get() # Still saving, with security note
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config_data, f, indent=4)
            if hasattr(self, 'log_text'): # Check if log_text is initialized
                self.log_message("Settings saved.")
        except IOError as e:
            if hasattr(self, 'log_text'):
                self.log_message(f"Error saving settings: {e}", error=True)
        except AttributeError:
            # log_text might not be available if called too early
            print("Log widget not available yet for 'Settings saved' message.")


    def load_config(self):
        """Loads settings from the config file."""
        self.default_email_body = ( # Define default before trying to load
            "Dear Hiring Manager at {{COMPANY_NAME}},\n\n"
            "I am writing to express my keen interest in an internship opportunity, potentially in a {{ROLE}} capacity or a related field, at {{COMPANY_NAME}}.\n\n"
            "My name is {{FIRST_NAME}} {{LAST_NAME}}, and I am a highly motivated student with a passion for [Your Field/Area of Interest]. "
            "I have been consistently impressed by {{COMPANY_NAME}}'s work in [Mention something specific about the company if possible, otherwise a general positive remark].\n\n"
            "I have attached my CV for your review, which further details my qualifications and experiences.\n\n"
            "Thank you for your time and consideration. I look forward to hearing from you soon.\n\n"
            "Sincerely,\n"
            "{{FIRST_NAME}} {{LAST_NAME}}\n"
            "[Your Phone Number - Optional]\n"
            "[Your LinkedIn Profile URL - Optional]")
        self.email_subject_var.set("Internship Application: {{ROLE}} at {{COMPANY_NAME}}")


        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f:
                    config_data = json.load(f)
                
                self.csv_file_path.set(config_data.get("csv_file_path", ""))
                self.cv_file_path.set(config_data.get("cv_file_path", ""))
                self.email_column_var.set(config_data.get("email_column", ""))
                
                loaded_mappings = config_data.get("column_mappings", {})
                for key, var in self.column_mappings.items():
                    var.set(loaded_mappings.get(key, "")) # Default to "" if key missing
                
                self.email_subject_var.set(config_data.get("email_subject", "Internship Application: {{ROLE}} at {{COMPANY_NAME}}"))
                self.default_email_body = config_data.get("email_body", self.default_email_body) # Use loaded or pre-defined default
                
                self.smtp_email_var.set(config_data.get("smtp_email", ""))
                self.smtp_password_var.set(config_data.get("smtp_password", ""))

        except (IOError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load or parse {CONFIG_FILE}: {e}. Using defaults.")
            # Defaults are already set prior to this block


    def create_widgets(self):
        """Creates all GUI widgets using a tabbed interface."""
        
        # --- Main container for tabs and log ---
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Notebook for tabs ---
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Tab 1: Setup & Mapping ---
        tab1 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab1, text='Setup & Mapping')
        self.create_tab1_widgets(tab1)

        # --- Tab 2: Email Content ---
        tab2 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab2, text='Email Content')
        self.create_tab2_widgets(tab2)

        # --- Tab 3: Send & Settings ---
        tab3 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab3, text='Send & Settings')
        self.create_tab3_widgets(tab3)
        
        # --- Log Frame (below tabs) ---
        log_frame = ttk.LabelFrame(main_container, text="Log", padding=10)
        log_frame.pack(fill=tk.X, padx=5, pady=(10,5), side=tk.BOTTOM)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=6, state='disabled')
        self.log_text.pack(fill=tk.X, expand=False) # Don't expand vertically, only X

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


    def create_tab1_widgets(self, parent_tab):
        """Creates widgets for Tab 1: Setup & Mapping."""
        # File Selection Frame
        file_frame = ttk.LabelFrame(parent_tab, text="1. Load Data & CV", padding=10)
        file_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        ttk.Button(file_frame, text="Load CSV File", command=self.load_csv_file).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(file_frame, textvariable=self.csv_file_path, wraplength=350).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Button(file_frame, text="Select CV (PDF)", command=self.select_cv_file).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(file_frame, textvariable=self.cv_file_path, wraplength=350).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        file_frame.columnconfigure(1, weight=1)

        # Column Mapping Frame
        mapping_frame = ttk.LabelFrame(parent_tab, text="2. Map CSV Columns", padding=10)
        mapping_frame.grid(row=1, column=0, padx=5, pady=10, sticky="nsew")
        
        ttk.Label(mapping_frame, text="CSV Column for Email Address (Required):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_column_menu = ttk.OptionMenu(mapping_frame, self.email_column_var, "Select Email Column", *[""])
        self.email_column_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(mapping_frame, text="--- Map Placeholders to CSV Columns (Auto-detected where possible): ---").grid(row=1, column=0, columnspan=2, pady=(10,2))
        
        self.placeholder_menus = {}
        current_row = 2
        for key, placeholder_text in DEFAULT_PLACEHOLDERS.items():
            label_text = f"{key.replace('_', ' ').title()} ({placeholder_text}):"
            ttk.Label(mapping_frame, text=label_text).grid(row=current_row, column=0, padx=5, pady=3, sticky="w")
            
            var = self.column_mappings[key]
            initial_val = var.get() if var.get() else ("Not Mapped" if not self.csv_headers else self.csv_headers[0])
            
            menu = ttk.OptionMenu(mapping_frame, var, initial_val, *self.csv_headers if self.csv_headers else ["Not Mapped"])
            menu.grid(row=current_row, column=1, padx=5, pady=3, sticky="ew")
            self.placeholder_menus[key] = menu
            current_row += 1
        
        mapping_frame.columnconfigure(1, weight=1)
        parent_tab.columnconfigure(0, weight=1)
        parent_tab.rowconfigure(1, weight=1) # Allow mapping_frame to expand


    def create_tab2_widgets(self, parent_tab):
        """Creates widgets for Tab 2: Email Content."""
        email_template_frame = ttk.LabelFrame(parent_tab, text="Email Template Editor", padding=10)
        email_template_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(email_template_frame, text="Subject:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(email_template_frame, textvariable=self.email_subject_var, width=80).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(email_template_frame, text="Body:").grid(row=1, column=0, padx=5, pady=5, sticky="nw")
        self.email_body_text_widget = scrolledtext.ScrolledText(email_template_frame, wrap=tk.WORD, height=15, width=80)
        self.email_body_text_widget.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        if hasattr(self, 'default_email_body'): # Ensure default_email_body is set during load_config
            self.email_body_text_widget.insert(tk.END, self.default_email_body)
        
        ttk.Button(email_template_frame, text="Preview Email (First Row)", command=self.preview_email).grid(row=2, column=1, padx=5, pady=10, sticky="e")
        
        email_template_frame.columnconfigure(1, weight=1)
        email_template_frame.rowconfigure(1, weight=1)


    def create_tab3_widgets(self, parent_tab):
        """Creates widgets for Tab 3: Send & Settings."""
        # SMTP Settings Frame
        smtp_frame = ttk.LabelFrame(parent_tab, text="Gmail SMTP Settings", padding=10)
        smtp_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        ttk.Label(smtp_frame, text="Your Gmail Address:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(smtp_frame, textvariable=self.smtp_email_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(smtp_frame, text="Gmail App Password:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(smtp_frame, textvariable=self.smtp_password_var, show="*", width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        smtp_frame.columnconfigure(1, weight=1)

        # Action Frame
        action_frame = ttk.LabelFrame(parent_tab, text="Actions", padding=10)
        action_frame.grid(row=0, column=1, padx=15, pady=5, sticky="ewns", rowspan=2)

        self.send_button = ttk.Button(action_frame, text="Send Emails", command=self.send_emails_process, style="Accent.TButton")
        self.send_button.pack(pady=10, fill=tk.X, ipady=5) # Make button a bit taller
        ttk.Button(action_frame, text="Save All Settings", command=self.save_config).pack(pady=5, fill=tk.X)
        
        parent_tab.columnconfigure(0, weight=1) # SMTP frame can expand
        parent_tab.columnconfigure(1, weight=0) # Action frame fixed width relative to its content

        # Add style for Accent button (if theme supports it, or use custom)
        style = ttk.Style()
        style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'), foreground="white")
        # Basic theming for some themes like 'clam' might allow background, others not easily
        try:
            style.map("Accent.TButton", background=[('active', 'green'), ('!disabled', 'darkgreen')])
        except tk.TclError:
            print("Warning: Theme does not support background mapping for Accent.TButton.")


    def on_closing(self):
        """Handle window close event."""
        if messagebox.askokcancel("Quit", "Do you want to save settings before quitting?"):
            self.save_config()
        self.root.destroy()

    def log_message(self, message, error=False):
        """Appends a message to the log text area."""
        if not hasattr(self, 'log_text') or self.log_text is None: # Check if log_text is initialized
            print(f"LOG ({'ERROR' if error else 'INFO'}): {message}")
            return

        self.log_text.config(state='normal')
        if error:
            self.log_text.insert(tk.END, f"ERROR: {message}\n", "error")
            self.log_text.tag_config("error", foreground="red")
        else:
            self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        if self.root: # Ensure root window exists
             self.root.update_idletasks()

    def _auto_detect_columns(self):
        """Attempts to auto-detect and set common columns like email, first name, company."""
        if not self.csv_headers:
            return

        # Auto-detect Email Column
        detected_email_col = self.email_column_var.get() # Preserve if already set (e.g. from config)
        if not detected_email_col or detected_email_col == "Not Mapped":
            for header in self.csv_headers:
                if header.lower().replace(" ", "").replace("_", "") in AUTO_DETECT_PATTERNS["email_column"]:
                    self.email_column_var.set(header)
                    self.log_message(f"Auto-detected Email column: '{header}'")
                    break
        
        # Auto-detect other placeholders
        for key, patterns in AUTO_DETECT_PATTERNS.items():
            if key == "email_column": continue # Already handled

            # Only auto-detect if not already set from config or manually
            current_mapping = self.column_mappings[key].get()
            if not current_mapping or current_mapping == "Not Mapped":
                for header in self.csv_headers:
                    if header.lower().replace(" ", "").replace("_", "") in patterns:
                        self.column_mappings[key].set(header)
                        self.log_message(f"Auto-detected {key.replace('_',' ').title()} column: '{header}'")
                        break # Found one, move to next key

    def _load_csv_data(self, file_path, silent=False):
        """Internal helper to load CSV data, update headers, and attempt auto-detection."""
        try:
            self.csv_data = []
            self.csv_headers = []
            with open(file_path, mode='r', encoding='utf-8-sig', newline='') as file:
                reader = csv.DictReader(file)
                if not reader.fieldnames:
                    if not silent:
                        messagebox.showerror("CSV Error", "CSV file is empty or has no headers.")
                    self.csv_file_path.set("")
                    return False
                self.csv_headers = reader.fieldnames
                for row in reader:
                    self.csv_data.append(row)
            
            if not self.csv_data and not silent:
                messagebox.showwarning("CSV Warning", "CSV file has headers but no data rows.")
            
            if not silent:
                self.log_message(f"Loaded {len(self.csv_data)} rows from {os.path.basename(file_path)}.")
            
            self._auto_detect_columns() # Attempt to auto-detect after loading headers
            self.update_column_mapping_dropdowns() # This will now use potentially auto-detected values
            return True

        except Exception as e:
            if not silent:
                messagebox.showerror("CSV Error", f"Failed to load CSV: {e}")
                self.log_message(f"Failed to load CSV: {e}", error=True)
            self.csv_file_path.set("")
            self.csv_headers = []
            self.csv_data = []
            self.update_column_mapping_dropdowns()
            return False


    def load_csv_file(self):
        """Opens a dialog to select a CSV file and loads its data."""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if file_path:
            self.csv_file_path.set(file_path)
            self._load_csv_data(file_path)


    def select_cv_file(self):
        """Opens a dialog to select a CV (PDF) file."""
        file_path = filedialog.askopenfilename(
            title="Select CV (PDF File)",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if file_path:
            if file_path.lower().endswith(".pdf"):
                self.cv_file_path.set(file_path)
                self.log_message(f"CV selected: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("File Error", "Please select a PDF file for the CV.")


    def update_column_mapping_dropdowns(self):
        """Updates the OptionMenu widgets for column mapping with current CSV headers."""
        options = ["Not Mapped"] + (self.csv_headers if self.csv_headers else [])
        
        # Update Email Column Dropdown
        current_email_col = self.email_column_var.get()
        self.email_column_menu['menu'].delete(0, 'end')
        for option_val in options:
            self.email_column_menu['menu'].add_command(label=option_val, command=tk._setit(self.email_column_var, option_val))
        
        if current_email_col in options:
            self.email_column_var.set(current_email_col)
        elif options and options[0] != "Not Mapped" and self.email_column_var.get() == "Select Email Column": # Default if nothing set
             self.email_column_var.set(options[0]) # Default to "Not Mapped" or first actual header
        elif not options:
            self.email_column_var.set("")


        # Update Placeholder Mapping Dropdowns
        for key, menu in self.placeholder_menus.items():
            current_selection = self.column_mappings[key].get()
            menu['menu'].delete(0, 'end')
            for option_val in options:
                menu['menu'].add_command(label=option_val, command=tk._setit(self.column_mappings[key], option_val))
            
            if current_selection in options: # Preserve selection if valid
                self.column_mappings[key].set(current_selection)
            elif self.column_mappings[key].get() == "" or self.column_mappings[key].get() not in options : # If current is invalid or empty
                if options:
                    self.column_mappings[key].set(options[0]) # Default to "Not Mapped"
                else:
                    self.column_mappings[key].set("")
        
        self.update_column_mapping_dropdowns_state()


    def update_column_mapping_dropdowns_state(self):
        """Enable/disable column mapping dropdowns based on CSV load status."""
        state = tk.NORMAL if self.csv_headers else tk.DISABLED
        if hasattr(self, 'email_column_menu'): # Ensure widget exists
            self.email_column_menu.config(state=state)
        if hasattr(self, 'placeholder_menus'):
            for menu in self.placeholder_menus.values():
                menu.config(state=state)

    def _is_valid_email(self, email_string):
        """Basic email validation using regex."""
        if not email_string or not isinstance(email_string, str):
            return False
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        return re.match(pattern, email_string) is not None

    def preview_email(self):
        """Shows a preview of the email using data from the first CSV row."""
        if not self.csv_data:
            messagebox.showinfo("Preview Info", "Load a CSV file with data to preview.")
            return
        if self.email_body_text_widget is None:
            messagebox.showerror("Error", "Email body editor not available.")
            return

        first_row = self.csv_data[0]
        subject = self.email_subject_var.get()
        body = self.email_body_text_widget.get("1.0", tk.END)

        preview_data = {}
        for key, placeholder in DEFAULT_PLACEHOLDERS.items():
            csv_column_name = self.column_mappings[key].get()
            if csv_column_name and csv_column_name != "Not Mapped" and csv_column_name in first_row:
                preview_data[placeholder] = first_row[csv_column_name]
            else:
                preview_data[placeholder] = f"[{key.upper()}_DATA_MISSING_OR_NOT_MAPPED]"

        for placeholder, value in preview_data.items():
            subject = subject.replace(placeholder, str(value)) # Ensure value is string
            body = body.replace(placeholder, str(value))
        
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Email Preview (First Row)")
        preview_window.geometry("600x450") # Slightly taller for better body preview
        preview_window.transient(self.root) # Keep on top of main window
        preview_window.grab_set() # Modal behavior

        style = ttk.Style(preview_window)
        style.configure("Preview.TLabel", font=("Helvetica", 10))
        style.configure("PreviewSubject.TLabel", font=("Helvetica", 11, "bold"))

        ttk.Label(preview_window, text="Subject:", style="PreviewSubject.TLabel").pack(pady=(10,2), anchor="w", padx=10)
        ttk.Label(preview_window, text=subject, wraplength=580, style="Preview.TLabel").pack(pady=(0,10), anchor="w", padx=10)
        
        ttk.Separator(preview_window, orient='horizontal').pack(fill='x', padx=10, pady=5)

        ttk.Label(preview_window, text="Body:", style="PreviewSubject.TLabel").pack(pady=(5,2), anchor="w", padx=10)
        body_preview_text = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, height=15, relief=tk.SOLID, borderwidth=1)
        body_preview_text.insert(tk.END, body)
        body_preview_text.config(state='disabled', font=("Helvetica", 10))
        body_preview_text.pack(pady=(0,10), padx=10, fill=tk.BOTH, expand=True)
        
        ttk.Button(preview_window, text="Close", command=preview_window.destroy).pack(pady=10)


    def send_emails_process(self):
        """Validates inputs and starts the email sending process."""
        if not self.csv_data: messagebox.showerror("Error", "No CSV data loaded."); return
        cv_path = self.cv_file_path.get()
        if not cv_path or not os.path.exists(cv_path): messagebox.showerror("Error", "CV file not selected or not found."); return
        if not cv_path.lower().endswith(".pdf"): messagebox.showerror("Error", "CV must be a PDF file."); return

        email_col_name = self.email_column_var.get()
        if not email_col_name or email_col_name == "Not Mapped" or email_col_name not in self.csv_headers:
            messagebox.showerror("Error", "Email column not selected or invalid. Please check Tab 1."); return

        sender_email = self.smtp_email_var.get()
        sender_password = self.smtp_password_var.get()
        if not sender_email or not sender_password: messagebox.showerror("Error", "Gmail address or App Password not provided."); return
        if not self._is_valid_email(sender_email): messagebox.showerror("Error", "Invalid sender Gmail address format."); return

        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body cannot be empty."); return

        if not messagebox.askyesno("Confirm Send", f"Are you sure you want to send emails to {len(self.csv_data)} recipients based on current mappings?"):
            return

        self.log_message(f"Starting email sending process for {len(self.csv_data)} potential recipients...")
        self.send_button.config(state=tk.DISABLED)

        sent_count = 0
        failed_count = 0

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, sender_password)
            self.log_message("Logged into Gmail SMTP server successfully.")

            for i, row_data in enumerate(self.csv_data):
                self.log_message(f"Processing row {i+1}...")
                recipient_email = row_data.get(email_col_name)
                if not recipient_email:
                    self.log_message(f"Skipped row {i+1}: Email address missing in column '{email_col_name}'.", error=True)
                    failed_count +=1; continue
                if not self._is_valid_email(recipient_email):
                    self.log_message(f"Skipped row {i+1}: Invalid recipient email format '{recipient_email}'.", error=True)
                    failed_count +=1; continue
                
                current_subject = subject_template
                current_body = body_template

                for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                    csv_col_for_placeholder = self.column_mappings[key].get()
                    value_to_insert = "" # Default to empty if not mapped or no data
                    if csv_col_for_placeholder and csv_col_for_placeholder != "Not Mapped" and csv_col_for_placeholder in row_data:
                        value_to_insert = row_data[csv_col_for_placeholder]
                    
                    current_subject = current_subject.replace(placeholder, str(value_to_insert))
                    current_body = current_body.replace(placeholder, str(value_to_insert))

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = current_subject
                msg.attach(MIMEText(current_body, 'plain', 'utf-8')) # Specify utf-8

                try:
                    with open(cv_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(cv_path)}")
                    msg.attach(part)
                except Exception as e:
                    self.log_message(f"Failed to attach CV for {recipient_email} (Row {i+1}): {e}", error=True)
                    failed_count += 1; continue

                try:
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                    self.log_message(f"Email sent to {recipient_email} (Row {i+1})")
                    sent_count += 1
                except Exception as e:
                    self.log_message(f"Failed to send email to {recipient_email} (Row {i+1}): {e}", error=True)
                    failed_count += 1
                
                self.root.update_idletasks()

            server.quit()
            self.log_message("Disconnected from SMTP server.")

        except smtplib.SMTPAuthenticationError:
            err_msg = "SMTP Authentication Error: Check Gmail address & App Password. Use App Password if 2FA is ON."
            self.log_message(err_msg, error=True)
            messagebox.showerror("SMTP Auth Error", err_msg)
            failed_count = len(self.csv_data) - sent_count
        except smtplib.SMTPConnectError:
            err_msg = "SMTP Connection Error: Could not connect to smtp.gmail.com. Check internet."
            self.log_message(err_msg, error=True)
            messagebox.showerror("SMTP Connection Error", err_msg)
            failed_count = len(self.csv_data) - sent_count
        except Exception as e:
            self.log_message(f"An unexpected error occurred: {e}", error=True)
            messagebox.showerror("Sending Error", f"An unexpected error: {e}")
        finally:
            self.log_message(f"Email sending process finished. Sent: {sent_count}, Failed: {failed_count}.")
            if hasattr(self, 'send_button'): self.send_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        # Apply a theme for a more modern look (optional, 'clam' is widely available)
        try:
            style = ttk.Style(root)
            available_themes = style.theme_names()
            # print("Available themes:", available_themes) # For debugging
            if 'clam' in available_themes:
                style.theme_use('clam')
            elif 'vista' in available_themes and os.name == 'nt': # Windows
                 style.theme_use('vista')
            elif 'aqua' in available_themes and os.name == 'posix': # macOS
                 style.theme_use('aqua')
        except Exception as e:
            print(f"Could not set theme: {e}")

        app = BulkEmailerApp(root)
        root.mainloop()
    except tk.TclError as e:
        print(f"Tkinter TclError: {e}")
        print("This Python script requires a graphical display environment to run.")
        print("Ensure you are running this on a system with a display (like a desktop environment),")
        print("or if using SSH, ensure X11 forwarding is enabled.")
    except Exception as e:
        print(f"An unexpected error occurred when starting the application: {e}")
        import traceback
        traceback.print_exc()

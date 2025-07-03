import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os
import re # For email validation
import datetime
import time # For progress bar updates and basic scheduling
import uuid # For unique campaign IDs

CONFIG_FILE = "bulk_emailer_config_profiles.json"
SCHEDULED_CAMPAIGNS_FILE = "scheduled_campaigns.json" # For persistent scheduled jobs

DEFAULT_PLACEHOLDERS = {
    "FIRST_NAME": "{{FIRST_NAME}}",
    "LAST_NAME": "{{LAST_NAME}}", # Retained for template consistency
    "COMPANY_NAME": "{{COMPANY_NAME}}",
    "ROLE": "{{ROLE}}",
}

AUTO_DETECT_PATTERNS = {
    "email_column": ['email', 'e-mail', 'email address', 'e mail'],
    "FIRST_NAME": ['firstname', 'first name', 'given name', 'first', 'fname'],
    "LAST_NAME": ['lastname', 'last name', 'surname', 'lname'],
    "COMPANY_NAME": ['companyname', 'company name', 'organization', 'company', 'employer'],
    "ROLE": ['role', 'position', 'job title', 'applied for', 'job']
}

DEFAULT_PROFILE_NAME = "Default Profile"

class BulkEmailerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Bulk Emailer (Batch Custom & Profile Scheduling)")
        self.root.geometry("1000x800") # Wider for new tab

        self.profiles = {}
        self.active_profile_name = tk.StringVar()
        self.profile_keys_for_dropdown = []

        self.csv_file_paths = [] 
        self.cv_file_path = tk.StringVar() 
        self.csv_headers = [] 
        self.csv_data = [] 

        self.email_column_var = tk.StringVar()
        self.column_mappings = {key: tk.StringVar() for key in DEFAULT_PLACEHOLDERS}

        self.email_subject_var = tk.StringVar()
        self.email_body_text_widget = None

        self.smtp_email_var = tk.StringVar()
        self.smtp_password_var = tk.StringVar()

        # CC Feature Vars
        self.enable_cc_var = tk.BooleanVar()
        self.cc_email_var = tk.StringVar()

        # Scheduling - now part of profile, these vars reflect active profile's schedule
        self.profile_schedule_date_var = tk.StringVar() 
        self.profile_schedule_time_var = tk.StringVar()

        # For the main scheduling UI (can override profile's default for a specific send)
        self.ui_schedule_date_var = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.ui_schedule_time_var = tk.StringVar(value=(datetime.datetime.now() + datetime.timedelta(minutes=15)).strftime("%H:%M"))


        self.manual_email_var = tk.StringVar()
        self.manual_first_name_var = tk.StringVar()
        self.manual_company_name_var = tk.StringVar()
        self.manual_role_var = tk.StringVar() 

        self.scheduled_campaigns = self.load_scheduled_campaigns_from_file()
        self.custom_email_batch = [] # List to hold custom email dicts

        self.load_app_config() 
        self.create_widgets()
        
        if not self.profiles: 
            self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True)
        elif self.active_profile_name.get() not in self.profiles:
            if self.profile_keys_for_dropdown: self.active_profile_name.set(self.profile_keys_for_dropdown[0])
            else: self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True)
        
        self.load_profile_data(self.active_profile_name.get()) # This will now load profile's schedule too
        self.update_column_mapping_dropdowns_state()
        
        self.root.after(2000, self.check_for_pending_scheduled_jobs) 
        self.root.after(60000, self.periodic_schedule_check) 


    def get_default_profile_settings(self):
        return {
            "cv_file_path": "", "email_column": "",
            "column_mappings": {key: "" for key in DEFAULT_PLACEHOLDERS},
            "email_subject": "Internship Application: {{ROLE}} at {{COMPANY_NAME}}",
            "email_body": ("Dear Hiring Manager at {{COMPANY_NAME}},\n\n"
                           "I am writing to express my keen interest in an internship opportunity, potentially in a {{ROLE}} capacity or a related field, at {{COMPANY_NAME}}.\n\n"
                           "My name is {{FIRST_NAME}} {{LAST_NAME}}, and I am a highly motivated student with a passion for [Your Field/Area of Interest].\n\n"
                           "I have attached my CV for your review.\n\n"
                           "Thank you for your time and consideration.\n\n"
                           "Sincerely,\n"
                           "{{FIRST_NAME}} {{LAST_NAME}}"),
            "smtp_email": "", "smtp_password": "",
            "schedule_date": "", "schedule_time": "", # Profile's default schedule
            "enable_cc": False, "cc_email": "" # Added CC settings
        }

    def save_app_config(self):
        self.save_current_profile_data_to_object()
        app_config = {"active_profile_name": self.active_profile_name.get(), "profiles": self.profiles}
        try:
            with open(CONFIG_FILE, "w") as f: json.dump(app_config, f, indent=4)
            self.log_message("Application configuration (all profiles) saved.")
        except Exception as e: self.log_message(f"Error saving application configuration: {e}", error=True)

    def load_app_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f: app_config = json.load(f)
                self.active_profile_name.set(app_config.get("active_profile_name", DEFAULT_PROFILE_NAME))
                self.profiles = app_config.get("profiles", {})
                if not self.profiles: 
                    self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                    if not self.active_profile_name.get(): self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                self.profile_keys_for_dropdown = list(self.profiles.keys())
                if not self.profile_keys_for_dropdown: 
                     self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                     self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                     self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
            else: 
                self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
        except Exception as e:
            self.log_message(f"Error loading config or config corrupted: {e}. Creating default.", error=True)
            self.active_profile_name.set(DEFAULT_PROFILE_NAME)
            self.profiles = {DEFAULT_PROFILE_NAME: self.get_default_profile_settings()}
            self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]

    def save_current_profile_data_to_object(self):
        profile_name = self.active_profile_name.get()
        if not profile_name or profile_name not in self.profiles: return
        if profile_name not in self.profiles: self.profiles[profile_name] = self.get_default_profile_settings()
        
        current_profile_data = self.profiles[profile_name]
        current_profile_data["cv_file_path"] = self.cv_file_path.get()
        current_profile_data["email_column"] = self.email_column_var.get()
        current_profile_data["column_mappings"] = {key: var.get() for key, var in self.column_mappings.items()}
        current_profile_data["email_subject"] = self.email_subject_var.get()
        if self.email_body_text_widget: current_profile_data["email_body"] = self.email_body_text_widget.get("1.0", tk.END).strip()
        else: current_profile_data["email_body"] = self.profiles[profile_name].get("email_body","")
        current_profile_data["smtp_email"] = self.smtp_email_var.get()
        current_profile_data["smtp_password"] = self.smtp_password_var.get()
        current_profile_data["schedule_date"] = self.profile_schedule_date_var.get() # Save profile's schedule
        current_profile_data["schedule_time"] = self.profile_schedule_time_var.get()
        current_profile_data["enable_cc"] = self.enable_cc_var.get()
        current_profile_data["cc_email"] = self.cc_email_var.get()


    def load_profile_data(self, profile_name):
        if not profile_name or profile_name not in self.profiles:
            self.log_message(f"Profile '{profile_name}' not found. Cannot load.", error=True)
            if DEFAULT_PROFILE_NAME in self.profiles and profile_name != DEFAULT_PROFILE_NAME:
                self.active_profile_name.set(DEFAULT_PROFILE_NAME); self.load_profile_data(DEFAULT_PROFILE_NAME)
            return

        profile_data = self.profiles[profile_name]
        self.active_profile_name.set(profile_name) 
        self.cv_file_path.set(profile_data.get("cv_file_path", ""))
        self.email_column_var.set(profile_data.get("email_column", ""))
        loaded_mappings = profile_data.get("column_mappings", {})
        for key, var_tk in self.column_mappings.items(): var_tk.set(loaded_mappings.get(key, ""))
        self.email_subject_var.set(profile_data.get("email_subject", "Internship Application: {{ROLE}} at {{COMPANY_NAME}}"))
        if self.email_body_text_widget: 
            self.email_body_text_widget.delete("1.0", tk.END)
            self.email_body_text_widget.insert("1.0", profile_data.get("email_body", self.get_default_profile_settings()["email_body"]))
        self.smtp_email_var.set(profile_data.get("smtp_email", ""))
        self.smtp_password_var.set(profile_data.get("smtp_password", ""))
        
        # Load profile's schedule into its dedicated vars
        self.profile_schedule_date_var.set(profile_data.get("schedule_date", ""))
        self.profile_schedule_time_var.set(profile_data.get("schedule_time", ""))

        # Load CC settings
        self.enable_cc_var.set(profile_data.get("enable_cc", False))
        self.cc_email_var.set(profile_data.get("cc_email", ""))
        
        # Update UI scheduling fields from profile's defaults if they are empty, or keep current UI if user typed something
        if not self.ui_schedule_date_var.get() and self.profile_schedule_date_var.get():
            self.ui_schedule_date_var.set(self.profile_schedule_date_var.get())
        if not self.ui_schedule_time_var.get() and self.profile_schedule_time_var.get():
             self.ui_schedule_time_var.set(self.profile_schedule_time_var.get())


        self.update_column_mapping_dropdowns() 
        self.log_message(f"Profile '{profile_name}' loaded.")
        if hasattr(self, 'toggle_cc_entry'): self.toggle_cc_entry() # Update CC entry state after loading

    def on_profile_selected(self, event=None):
        selected_profile = self.active_profile_name.get()
        self.load_profile_data(selected_profile)

    def create_new_profile_dialog(self):
        profile_name = simpledialog.askstring("New Profile", "Enter name for the new profile:", parent=self.root)
        if profile_name:
            if profile_name in self.profiles: messagebox.showerror("Error", f"Profile '{profile_name}' already exists.")
            else: self.create_new_profile(profile_name, make_active=True)
    
    def create_new_profile(self, profile_name, make_active=False, initial_setup=False):
        new_profile_settings = self.get_default_profile_settings()
        if not initial_setup: 
            current_active_profile_name_for_inheritance = self.active_profile_name.get()
            if current_active_profile_name_for_inheritance and current_active_profile_name_for_inheritance in self.profiles:
                active_profile_data = self.profiles[current_active_profile_name_for_inheritance]
                new_profile_settings["smtp_email"] = active_profile_data.get("smtp_email", "")
                new_profile_settings["smtp_password"] = active_profile_data.get("smtp_password", "")
                new_profile_settings["schedule_date"] = active_profile_data.get("schedule_date", "") # Inherit schedule too
                new_profile_settings["schedule_time"] = active_profile_data.get("schedule_time", "")
                new_profile_settings["enable_cc"] = active_profile_data.get("enable_cc", False) # Inherit CC settings
                new_profile_settings["cc_email"] = active_profile_data.get("cc_email", "")
                self.log_message(f"New profile '{profile_name}' inherited SMTP, Schedule, and CC from '{current_active_profile_name_for_inheritance}'.")
        self.profiles[profile_name] = new_profile_settings
        self.profile_keys_for_dropdown = list(self.profiles.keys())
        self.update_profile_dropdown()
        if make_active:
            self.active_profile_name.set(profile_name); self.load_profile_data(profile_name)
        self.log_message(f"Profile '{profile_name}' created.")
        if not initial_setup: self.save_app_config() 

    def delete_current_profile_dialog(self):
        profile_name_to_delete = self.active_profile_name.get()
        if not profile_name_to_delete or profile_name_to_delete == DEFAULT_PROFILE_NAME:
            messagebox.showerror("Error", "Cannot delete the default profile or no profile selected."); return
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete profile '{profile_name_to_delete}'? This cannot be undone."):
            if profile_name_to_delete in self.profiles:
                del self.profiles[profile_name_to_delete]
                self.profile_keys_for_dropdown = list(self.profiles.keys())
                new_active = DEFAULT_PROFILE_NAME if DEFAULT_PROFILE_NAME in self.profiles else (self.profile_keys_for_dropdown[0] if self.profile_keys_for_dropdown else "")
                self.active_profile_name.set(new_active); self.update_profile_dropdown() 
                if new_active: self.load_profile_data(new_active)
                else: self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True)
                self.log_message(f"Profile '{profile_name_to_delete}' deleted."); self.save_app_config()

    def update_profile_dropdown(self):
        if not hasattr(self, 'profile_menu'): return 
        menu = self.profile_menu['menu']; menu.delete(0, 'end')
        if not self.profile_keys_for_dropdown and DEFAULT_PROFILE_NAME not in self.profiles:
            self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
            self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
            if not self.active_profile_name.get(): self.active_profile_name.set(DEFAULT_PROFILE_NAME)
        elif not self.profile_keys_for_dropdown and DEFAULT_PROFILE_NAME in self.profiles:
             self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
        for profile_key in self.profile_keys_for_dropdown:
            menu.add_command(label=profile_key, command=lambda pk=profile_key: self.set_and_load_profile(pk))
        current_active = self.active_profile_name.get()
        if current_active not in self.profile_keys_for_dropdown:
             if self.profile_keys_for_dropdown: self.active_profile_name.set(self.profile_keys_for_dropdown[0])

    def set_and_load_profile(self, profile_key):
        self.active_profile_name.set(profile_key); self.on_profile_selected()

    def create_widgets(self):
        main_container = ttk.Frame(self.root); main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.notebook = ttk.Notebook(main_container); self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tab_profile_csv = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_profile_csv, text='Profiles & CSV'); self.create_tab_profile_csv(tab_profile_csv)
        tab_mapping = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_mapping, text='Column Mapping'); self.create_tab_mapping(tab_mapping)
        tab_email_content = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_email_content, text='Email Content'); self.create_tab_email_content(tab_email_content)
        tab_custom_batch = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_custom_batch, text='Batch Custom Emails'); self.create_tab_custom_batch(tab_custom_batch)
        tab_manual_send = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_manual_send, text='Manual Send'); self.create_tab_manual_send(tab_manual_send)
        tab_settings_send = ttk.Frame(self.notebook, padding=10); self.notebook.add(tab_settings_send, text='Settings & Send'); self.create_tab_settings_send(tab_settings_send)
        
        log_frame = ttk.LabelFrame(main_container, text="Log", padding=10); log_frame.pack(fill=tk.X, padx=5, pady=(10,5), side=tk.BOTTOM)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=5, state='disabled', font=("Helvetica", 9)); self.log_text.pack(fill=tk.X, expand=False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_profile_dropdown() 

    def create_tab_profile_csv(self, parent_tab):
        profile_frame = ttk.LabelFrame(parent_tab, text="User Profiles", padding=10); profile_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(profile_frame, text="Active Profile:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        initial_active_profile = self.active_profile_name.get()
        if not initial_active_profile or initial_active_profile not in self.profile_keys_for_dropdown:
            initial_active_profile = self.profile_keys_for_dropdown[0] if self.profile_keys_for_dropdown else "None"
            self.active_profile_name.set(initial_active_profile)
        self.profile_menu = ttk.OptionMenu(profile_frame, self.active_profile_name, initial_active_profile, *(self.profile_keys_for_dropdown or ["None"]), command=self.on_profile_selected)
        self.profile_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(profile_frame, text="New Profile", command=self.create_new_profile_dialog).grid(row=1, column=0, padx=5, pady=2, sticky="ew")
        ttk.Button(profile_frame, text="Save Active Profile Settings", command=self.save_app_config).grid(row=1, column=1, padx=5, pady=2, sticky="ew") 
        ttk.Button(profile_frame, text="Delete Current Profile", command=self.delete_current_profile_dialog).grid(row=1, column=2, padx=5, pady=2, sticky="ew")
        profile_frame.columnconfigure(1, weight=1)
        file_frame = ttk.LabelFrame(parent_tab, text="Load Data & CV (for current session/profile)", padding=10); file_frame.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        ttk.Button(file_frame, text="Load CSV File(s)", command=self.load_csv_files).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.csv_paths_label = ttk.Label(file_frame, text="No CSVs loaded.", wraplength=350); self.csv_paths_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(file_frame, text="Select CV (PDF for active profile)", command=self.select_cv_file).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(file_frame, textvariable=self.cv_file_path, wraplength=350).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        file_frame.columnconfigure(1, weight=1); parent_tab.columnconfigure(0, weight=1)

    def create_tab_mapping(self, parent_tab):
        mapping_frame = ttk.LabelFrame(parent_tab, text="Map CSV Columns (for active profile)", padding=10); mapping_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(mapping_frame, text="CSV Column for Email Address (Required):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_column_menu = ttk.OptionMenu(mapping_frame, self.email_column_var, self.email_column_var.get() or "Select Email Column", *(self.csv_headers or ["Not Mapped"]))
        self.email_column_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(mapping_frame, text="--- Map Placeholders to CSV Columns (Auto-detected where possible): ---").grid(row=1, column=0, columnspan=2, pady=(10,2))
        self.placeholder_menus = {}; current_row = 2
        for key, placeholder_text in DEFAULT_PLACEHOLDERS.items():
            label_text = f"{key.replace('_', ' ').title()} ({placeholder_text}):"; ttk.Label(mapping_frame, text=label_text).grid(row=current_row, column=0, padx=5, pady=3, sticky="w")
            var = self.column_mappings[key]
            initial_val_map = var.get() if var.get() else ("Not Mapped" if not self.csv_headers else (self.csv_headers[0] if self.csv_headers else "Not Mapped"))
            menu = ttk.OptionMenu(mapping_frame, var, initial_val_map, *(self.csv_headers if self.csv_headers else ["Not Mapped"])); menu.grid(row=current_row, column=1, padx=5, pady=3, sticky="ew")
            self.placeholder_menus[key] = menu; current_row += 1
        mapping_frame.columnconfigure(1, weight=1)

    def create_tab_email_content(self, parent_tab):
        email_template_frame = ttk.LabelFrame(parent_tab, text="Email Template Editor (for active profile)", padding=10); email_template_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(email_template_frame, text="Subject:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(email_template_frame, textvariable=self.email_subject_var, width=80).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(email_template_frame, text="Body:").grid(row=1, column=0, padx=5, pady=5, sticky="nw")
        self.email_body_text_widget = scrolledtext.ScrolledText(email_template_frame, wrap=tk.WORD, height=15, width=80, font=("Helvetica", 10)); self.email_body_text_widget.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        ttk.Button(email_template_frame, text="Preview Email (using first CSV row if available)", command=self.preview_email).grid(row=2, column=1, padx=5, pady=10, sticky="e")
        email_template_frame.columnconfigure(1, weight=1); email_template_frame.rowconfigure(1, weight=1)

    def create_tab_custom_batch(self, parent_tab):
        main_custom_frame = ttk.Frame(parent_tab); main_custom_frame.pack(fill=tk.BOTH, expand=True)
        list_frame = ttk.LabelFrame(main_custom_frame, text="Custom Email Batch", padding=10); list_frame.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)
        self.custom_emails_listbox = tk.Listbox(list_frame, height=10, selectmode=tk.SINGLE); self.custom_emails_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        list_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.custom_emails_listbox.yview); list_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.custom_emails_listbox.config(yscrollcommand=list_scrollbar.set)
        list_actions_frame = ttk.Frame(list_frame); list_actions_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5,0))
        ttk.Button(list_actions_frame, text="Add Email", command=self.add_or_edit_custom_email_dialog).pack(pady=2, fill=tk.X)
        ttk.Button(list_actions_frame, text="Edit Selected", command=lambda: self.add_or_edit_custom_email_dialog(edit_mode=True)).pack(pady=2, fill=tk.X)
        ttk.Button(list_actions_frame, text="Remove Selected", command=self.remove_selected_custom_email).pack(pady=2, fill=tk.X)
        ttk.Button(list_actions_frame, text="Clear Batch", command=self.clear_custom_email_batch).pack(pady=(10,2), fill=tk.X)
        send_batch_frame = ttk.LabelFrame(main_custom_frame, text="Send Batch", padding=10); send_batch_frame.pack(pady=10, padx=5, fill=tk.X)
        self.custom_batch_send_button = ttk.Button(send_batch_frame, text="Send All Custom Emails in Batch", command=self.send_custom_email_batch_process, style="Accent.TButton"); self.custom_batch_send_button.pack(pady=5, ipady=4)
        self.custom_batch_progress_bar = ttk.Progressbar(send_batch_frame, orient="horizontal", length=300, mode="determinate"); self.custom_batch_progress_bar.pack(pady=5, fill=tk.X)
        self.refresh_custom_emails_listbox()

    def create_tab_manual_send(self, parent_tab):
        manual_frame = ttk.LabelFrame(parent_tab, text="Send Single Email Manually", padding=10); manual_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(manual_frame, text="Recipient Email:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(manual_frame, textvariable=self.manual_email_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(manual_frame, text="{{FIRST_NAME}}:").grid(row=1, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(manual_frame, textvariable=self.manual_first_name_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(manual_frame, text="{{COMPANY_NAME}}:").grid(row=2, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(manual_frame, textvariable=self.manual_company_name_var, width=50).grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(manual_frame, text="{{ROLE}} (Optional):").grid(row=3, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(manual_frame, textvariable=self.manual_role_var, width=50).grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        manual_frame.columnconfigure(1, weight=1)
        action_buttons_frame = ttk.Frame(manual_frame); action_buttons_frame.grid(row=4, column=0, columnspan=2, pady=15)
        ttk.Button(action_buttons_frame, text="Preview Manual Email", command=lambda: self.preview_email(manual_mode=True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_buttons_frame, text="Send Manual Email", command=self.send_manual_email_process, style="Accent.TButton").pack(side=tk.LEFT, padx=5)

    def create_tab_settings_send(self, parent_tab):
        smtp_frame = ttk.LabelFrame(parent_tab, text="Gmail SMTP Settings (for active profile)", padding=10); smtp_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(smtp_frame, text="Your Gmail Address:").grid(row=0, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(smtp_frame, textvariable=self.smtp_email_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Label(smtp_frame, text="Gmail App Password:").grid(row=1, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(smtp_frame, textvariable=self.smtp_password_var, show="*", width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        smtp_frame.columnconfigure(1, weight=1)

        # CC Frame
        cc_frame = ttk.LabelFrame(parent_tab, text="CC Options (for active profile)", padding=10)
        cc_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.cc_checkbutton = ttk.Checkbutton(cc_frame, text="Enable CC for all sent emails", variable=self.enable_cc_var, command=self.toggle_cc_entry)
        self.cc_checkbutton.grid(row=0, column=0, columnspan=2, padx=5, pady=2, sticky="w")
        ttk.Label(cc_frame, text="CC Email Address:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cc_entry = ttk.Entry(cc_frame, textvariable=self.cc_email_var, width=40)
        self.cc_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        cc_frame.columnconfigure(1, weight=1)
        self.toggle_cc_entry() # Set initial state

        # Scheduling Frame
        schedule_outer_frame = ttk.LabelFrame(parent_tab, text="Scheduling (Overrides Profile Default for this Send)", padding=10)
        schedule_outer_frame.grid(row=2, column=0, padx=5, pady=10, sticky="ew")
        schedule_frame = ttk.Frame(schedule_outer_frame); schedule_frame.pack(fill=tk.X)
        ttk.Label(schedule_frame, text="Schedule Date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(schedule_frame, textvariable=self.ui_schedule_date_var, width=12).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(schedule_frame, text="Schedule Time (HH:MM, 24-hr):").grid(row=0, column=2, padx=(10,5), pady=5, sticky="w")
        ttk.Entry(schedule_frame, textvariable=self.ui_schedule_time_var, width=8).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(schedule_frame, text="Active Profile Default: Date:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(schedule_frame, textvariable=self.profile_schedule_date_var).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(schedule_frame, text="Time:").grid(row=1, column=2, padx=(10,5), pady=2, sticky="w")
        ttk.Label(schedule_frame, textvariable=self.profile_schedule_time_var).grid(row=1, column=3, padx=5, pady=2, sticky="w")
        info_label = ttk.Label(schedule_frame, text="(Leave UI Date/Time blank to use profile's default, or fill to override for current bulk send.)", font=("Helvetica", 8), wraplength=400)
        info_label.grid(row=2, column=0, columnspan=4, padx=5, pady=2, sticky="w")

        # Action Frame
        action_frame = ttk.LabelFrame(parent_tab, text="Bulk Sending Actions", padding=10); action_frame.grid(row=0, column=1, padx=15, pady=5, sticky="nsew", rowspan=3)
        self.send_button = ttk.Button(action_frame, text="Send/Schedule Bulk Emails", command=self.send_emails_process, style="Accent.TButton"); self.send_button.pack(pady=10, fill=tk.X, ipady=4)
        ttk.Button(action_frame, text="Send Test Email to Myself", command=self.send_test_email_process).pack(pady=5, fill=tk.X)
        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", length=200, mode="determinate"); self.progress_bar.pack(pady=10, fill=tk.X)
        parent_tab.columnconfigure(0, weight=1); parent_tab.columnconfigure(1, weight=0)
        style = ttk.Style(); style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'))
        try: style.map("Accent.TButton", foreground=[('!disabled', 'white')], background=[('active', 'darkgreen'), ('!disabled', 'green')])
        except tk.TclError: self.log_message("Note: Theme may not fully support custom button styling.", error=False)

    def toggle_cc_entry(self):
        """Enables or disables the CC email entry based on the checkbutton state."""
        if hasattr(self, 'cc_entry'):
            state = "normal" if self.enable_cc_var.get() else "disabled"
            self.cc_entry.config(state=state)

    def on_closing(self):
        self.log_message("Closing application. Auto-saving all profiles and settings...")
        self.save_app_config(); self.save_scheduled_campaigns_to_file() 
        self.root.destroy()

    def log_message(self, message, error=False):
        if not hasattr(self, 'log_text') or self.log_text is None: print(f"LOG ({'ERROR' if error else 'INFO'}): {message}"); return
        try:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} - {message}\n", "error_tag" if error else "info_tag")
            self.log_text.tag_config("error_tag", foreground="red"); self.log_text.tag_config("info_tag", foreground="black")
            self.log_text.see(tk.END); self.log_text.config(state='disabled')
            if self.root and self.root.winfo_exists(): self.root.update_idletasks()
        except tk.TclError: print(f"LOG (TclError suppressed): {message}")

    def _auto_detect_columns(self): 
        if not self.csv_headers: return
        current_email_col_setting = self.email_column_var.get()
        if not current_email_col_setting or current_email_col_setting not in self.csv_headers:
            detected = False
            for header in self.csv_headers:
                if header.lower().replace(" ", "").replace("_", "") in AUTO_DETECT_PATTERNS["email_column"]:
                    self.email_column_var.set(header); self.log_message(f"Auto-detected Email column: '{header}'"); detected = True; break
        for key, patterns in AUTO_DETECT_PATTERNS.items():
            if key == "email_column": continue
            current_mapping = self.column_mappings[key].get()
            if not current_mapping or current_mapping == "Not Mapped" or current_mapping not in self.csv_headers:
                detected_placeholder = False
                for header in self.csv_headers:
                    normalized_header = header.lower().replace(" ", "").replace("_", "")
                    if normalized_header in patterns:
                        self.column_mappings[key].set(header); self.log_message(f"Auto-detected {key.replace('_',' ').title()} column: '{header}'"); detected_placeholder = True; break
                if not detected_placeholder and self.column_mappings[key].get() not in self.csv_headers: self.column_mappings[key].set("Not Mapped")
        self.update_column_mapping_dropdowns()

    def _load_csv_data_from_paths(self, file_paths, silent=False): 
        self.csv_data = []; combined_headers = set(); all_rows = []
        if not file_paths: self.csv_headers = []; self.csv_paths_label.config(text="No CSVs loaded."); self.update_column_mapping_dropdowns(); return True 
        for file_path in file_paths:
            try:
                with open(file_path, mode='r', encoding='utf-8-sig', newline='') as file:
                    reader = csv.DictReader(file)
                    if not reader.fieldnames:
                        if not silent: messagebox.showwarning("CSV Warning", f"CSV file '{os.path.basename(file_path)}' is empty or has no headers. Skipping.")
                        continue
                    current_file_rows = list(reader)
                    if not current_file_rows and not silent: messagebox.showwarning("CSV Warning", f"CSV file '{os.path.basename(file_path)}' has headers but no data rows.")
                    all_rows.extend(current_file_rows)
                    for header in reader.fieldnames: combined_headers.add(header)
                if not silent: self.log_message(f"Successfully processed {os.path.basename(file_path)}.")
            except Exception as e:
                if not silent: messagebox.showerror("CSV Error", f"Failed to load/parse {os.path.basename(file_path)}: {e}"); self.log_message(f"Failed to load {os.path.basename(file_path)}: {e}", error=True)
        self.csv_headers = sorted(list(combined_headers)); self.csv_data = all_rows
        if not self.csv_data and not silent and file_paths: self.log_message("Warning: All loaded CSVs combined resulted in no data rows.", error=False)
        elif self.csv_data: self.log_message(f"Total {len(self.csv_data)} rows loaded from {len(file_paths)} CSV file(s).")
        self.csv_paths_label.config(text=f"{len(file_paths)} CSV(s) loaded: " + ", ".join([os.path.basename(p) for p in file_paths]) if file_paths else "No CSVs loaded.")
        self._auto_detect_columns(); return True

    def load_csv_files(self): 
        filepaths = filedialog.askopenfilenames(title="Select CSV Files", filetypes=(("CSV files", "*.csv"), ("All files", "*.*")))
        if filepaths: self.csv_file_paths = list(filepaths); self._load_csv_data_from_paths(self.csv_file_paths)

    def select_cv_file(self): 
        file_path = filedialog.askopenfilename(title="Select CV (PDF File for Active Profile)", filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")))
        if file_path:
            if file_path.lower().endswith(".pdf"): self.cv_file_path.set(file_path); self.log_message(f"CV selected for current profile: {os.path.basename(file_path)}")
            else: messagebox.showerror("File Error", "Please select a PDF file for the CV.")

    def update_column_mapping_dropdowns(self): 
        options = ["Not Mapped"] + (self.csv_headers if self.csv_headers else [])
        if hasattr(self, 'email_column_menu'):
            current_email_col_val = self.email_column_var.get(); self.email_column_menu['menu'].delete(0, 'end')
            default_email_option = current_email_col_val if current_email_col_val in options else options[0]
            self.email_column_var.set(default_email_option) 
            for option_val in options: self.email_column_menu['menu'].add_command(label=option_val, command=tk._setit(self.email_column_var, option_val))
        if hasattr(self, 'placeholder_menus'):
            for key, menu_widget in self.placeholder_menus.items():
                current_placeholder_val = self.column_mappings[key].get(); menu_widget['menu'].delete(0, 'end')
                default_placeholder_option = current_placeholder_val if current_placeholder_val in options else options[0]
                self.column_mappings[key].set(default_placeholder_option) 
                for option_val in options: menu_widget['menu'].add_command(label=option_val, command=tk._setit(self.column_mappings[key], option_val))
        self.update_column_mapping_dropdowns_state()

    def update_column_mapping_dropdowns_state(self): 
        state = tk.NORMAL if self.csv_headers else tk.DISABLED
        if hasattr(self, 'email_column_menu'): self.email_column_menu.config(state=state)
        if hasattr(self, 'placeholder_menus'):
            for menu in self.placeholder_menus.values(): menu.config(state=state)

    def _is_valid_email(self, email_string): 
        if not email_string or not isinstance(email_string, str): return False
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"; return re.match(pattern, email_string) is not None

    def _validate_schedule_datetime(self, date_str, time_str):
        if not date_str and not time_str: return None 
        if not date_str or not time_str: return "Invalid" 
        try:
            dt_str = f"{date_str} {time_str}"
            return datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
            return "Invalid"

    def preview_email(self, manual_mode=False, custom_email_data=None): 
        if self.email_body_text_widget is None: messagebox.showerror("Error", "Email body editor not available."); return
        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END)
        preview_fill_data = {}
        if custom_email_data: 
            if custom_email_data["use_template"]:
                for key, placeholder_text in DEFAULT_PLACEHOLDERS.items():
                    preview_fill_data[placeholder_text] = custom_email_data["template_placeholders"].get(key, f"[{key}_MISSING]")
            else: 
                subject_template = custom_email_data["subject"]
                body_template = custom_email_data["body"]
        elif manual_mode:
            preview_fill_data[DEFAULT_PLACEHOLDERS["FIRST_NAME"]] = self.manual_first_name_var.get() or "[MANUAL_FIRST_NAME]"
            preview_fill_data[DEFAULT_PLACEHOLDERS["LAST_NAME"]] = "" 
            preview_fill_data[DEFAULT_PLACEHOLDERS["COMPANY_NAME"]] = self.manual_company_name_var.get() or "[MANUAL_COMPANY_NAME]"
            preview_fill_data[DEFAULT_PLACEHOLDERS["ROLE"]] = self.manual_role_var.get() or "[MANUAL_ROLE]"
        else: 
            if not self.csv_data: messagebox.showinfo("Preview Info", "Load CSV data to preview bulk email."); return
            first_row = self.csv_data[0]
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                csv_col_name = self.column_mappings[key].get()
                if csv_col_name and csv_col_name != "Not Mapped" and csv_col_name in first_row:
                    preview_fill_data[placeholder] = first_row[csv_col_name]
                else: preview_fill_data[placeholder] = f"[{key.upper()}_DATA]"
        final_subject = subject_template; final_body = body_template
        if not custom_email_data or custom_email_data["use_template"]: 
            for placeholder, value in preview_fill_data.items():
                final_subject = final_subject.replace(placeholder, str(value))
                final_body = final_body.replace(placeholder, str(value))
        preview_window = tk.Toplevel(self.root); preview_window.title("Email Preview"); preview_window.geometry("600x450")
        preview_window.transient(self.root); preview_window.grab_set()
        ttk.Label(preview_window, text="Subject:", font=("Helvetica", 11, "bold")).pack(pady=(10,2), anchor="w", padx=10)
        ttk.Label(preview_window, text=final_subject, wraplength=580, font=("Helvetica", 10)).pack(pady=(0,10), anchor="w", padx=10)
        ttk.Separator(preview_window, orient='horizontal').pack(fill='x', padx=10, pady=5)
        ttk.Label(preview_window, text="Body:", font=("Helvetica", 11, "bold")).pack(pady=(5,2), anchor="w", padx=10)
        body_prev_text = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, height=15, relief=tk.SOLID, borderwidth=1, font=("Helvetica", 10))
        body_prev_text.insert(tk.END, final_body); body_prev_text.config(state='disabled')
        body_prev_text.pack(pady=(0,10), padx=10, fill=tk.BOTH, expand=True)
        ttk.Button(preview_window, text="Close", command=preview_window.destroy).pack(pady=10)

    def _perform_email_sending(self, emails_to_send_list, is_test=False, is_custom_batch=False):
        cv_path = self.cv_file_path.get(); sender_email = self.smtp_email_var.get(); sender_password = self.smtp_password_var.get()
        enable_cc = self.enable_cc_var.get()
        cc_email = self.cc_email_var.get()

        if not is_test and cv_path and not os.path.exists(cv_path):
             self.log_message(f"CV path '{cv_path}' is invalid. Sending without CV.", error=True); cv_path = None
        elif not is_test and cv_path and not cv_path.lower().endswith(".pdf"):
             self.log_message(f"CV file '{cv_path}' is not a PDF. Sending without CV.", error=True); cv_path = None
        
        progress_bar_to_use = self.custom_batch_progress_bar if is_custom_batch else self.progress_bar
        send_button_to_use = self.custom_batch_send_button if is_custom_batch else self.send_button

        self.log_message(f"Starting SMTP process for {len(emails_to_send_list)} email(s)...")
        if hasattr(send_button_to_use, 'config'): send_button_to_use.config(state=tk.DISABLED)
        if hasattr(progress_bar_to_use, 'config'): progress_bar_to_use['value'] = 0; progress_bar_to_use['maximum'] = len(emails_to_send_list) if emails_to_send_list else 1
        sent_count = 0; failed_count = 0
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587); server.ehlo(); server.starttls(); server.ehlo(); server.login(sender_email, sender_password)
            self.log_message("Logged into Gmail SMTP server.")
            for i, email_details in enumerate(emails_to_send_list):
                recipient_email = email_details['recipient_email']; current_subject = email_details['subject']; current_body = email_details['body']
                row_identifier = email_details.get('row_identifier', f"item {i+1}")
                
                msg = MIMEMultipart(); msg['From'] = sender_email; msg['To'] = recipient_email; msg['Subject'] = current_subject
                
                all_recipients = [recipient_email]
                if enable_cc and cc_email and self._is_valid_email(cc_email):
                    msg['Cc'] = cc_email
                    all_recipients.append(cc_email)

                msg.attach(MIMEText(current_body, 'plain', 'utf-8'))

                if cv_path and os.path.exists(cv_path) and cv_path.lower().endswith(".pdf"): 
                    try:
                        with open(cv_path, "rb") as attachment_file: part = MIMEBase('application', 'octet-stream'); part.set_payload(attachment_file.read())
                        encoders.encode_base64(part); part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(cv_path)}"); msg.attach(part)
                    except Exception as e:
                        self.log_message(f"Failed to attach CV for {recipient_email} ({row_identifier}): {e}", error=True)
                        if not is_test: failed_count += 1; self.update_progress(i + 1, is_custom_batch); continue 
                try:
                    server.sendmail(sender_email, all_recipients, msg.as_string()) # Use all_recipients here
                    self.log_message(f"Email sent to {recipient_email} ({row_identifier})"); sent_count += 1
                except Exception as e:
                    self.log_message(f"Failed to send email to {recipient_email} ({row_identifier}): {e}", error=True)
                    if not is_test: failed_count += 1
                if not is_test: self.update_progress(i + 1, is_custom_batch)
                time.sleep(0.05) 
            server.quit(); self.log_message("Disconnected from SMTP server.")
        except smtplib.SMTPAuthenticationError: err = "SMTP Auth Error. Check Gmail & App Password."; self.log_message(err, error=True); messagebox.showerror("SMTP Auth Error", err);
        except smtplib.SMTPConnectError: err = "SMTP Connection Error. Check internet."; self.log_message(err, error=True); messagebox.showerror("SMTP Connection Error", err);
        except Exception as e: self.log_message(f"Unexpected error during sending: {e}", error=True); messagebox.showerror("Sending Error", f"Unexpected error: {e}")
        finally:
            self.log_message(f"Process finished. Sent: {sent_count}, Failed: {failed_count if not is_test else 'N/A for test'}.")
            if hasattr(send_button_to_use, 'config'): send_button_to_use.config(state=tk.NORMAL)
            if hasattr(progress_bar_to_use, 'config') and not is_test and emails_to_send_list : progress_bar_to_use['value'] = progress_bar_to_use['maximum'] 

    def update_progress(self, current_step, is_custom_batch=False):
        progress_bar_to_use = self.custom_batch_progress_bar if is_custom_batch else self.progress_bar
        if hasattr(progress_bar_to_use, 'config'):
            progress_bar_to_use['value'] = current_step
            if self.root and self.root.winfo_exists(): self.root.update_idletasks()

    def send_emails_process(self, campaign_id_to_send=None, scheduled_campaign_data=None):
        emails_to_send_list = []
        sender_email_for_campaign = self.smtp_email_var.get()
        sender_password_for_campaign = self.smtp_password_var.get()
        cv_path_for_campaign = self.cv_file_path.get()
        # CC settings for campaign are read from UI/profile inside _perform_email_sending

        if scheduled_campaign_data:
            self.log_message(f"Processing pre-prepared scheduled campaign ID: {campaign_id_to_send}")
            emails_to_send_list = scheduled_campaign_data["emails_to_send_list"]
            # We will use the SMTP/CV settings from the *currently active profile* when sending a scheduled job,
            # but the email list is from the time of scheduling. This is a design choice for simplicity.
            # An alternative would be to store all profile settings with the job.
        else: 
            if not self.csv_data: messagebox.showerror("Error", "No CSV data loaded."); return
            email_col_name = self.email_column_var.get()
            if not email_col_name or email_col_name == "Not Mapped" or email_col_name not in self.csv_headers:
                messagebox.showerror("Error", "Email column not selected/invalid. Check 'Column Mapping' tab."); return
            if not sender_email_for_campaign or not sender_password_for_campaign: messagebox.showerror("Error", "Gmail address or App Password not set in active profile."); return
            if not self._is_valid_email(sender_email_for_campaign): messagebox.showerror("Error", "Invalid sender Gmail address format."); return
            subject_template = self.email_subject_var.get()
            body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
            if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return
            if cv_path_for_campaign and not os.path.exists(cv_path_for_campaign):
                if not messagebox.askyesno("CV Path Invalid", f"CV path invalid:\n'{cv_path_for_campaign}'\nContinue without CV?"): return
                cv_path_for_campaign = None 
            elif cv_path_for_campaign and not cv_path_for_campaign.lower().endswith(".pdf"):
                messagebox.showerror("Error", "CV file must be a PDF."); return
            elif not cv_path_for_campaign: self.log_message("No CV selected. Emails will be sent without attachments.", error=False)

            for i, row_data in enumerate(self.csv_data):
                recipient_email = row_data.get(email_col_name)
                if not recipient_email or not self._is_valid_email(recipient_email):
                    self.log_message(f"Skipping row {i+1}: Invalid/missing email '{recipient_email}'.", error=True); continue
                current_subject = subject_template; current_body = body_template
                for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                    csv_col_for_placeholder = self.column_mappings[key].get(); value_to_insert = ""
                    if csv_col_for_placeholder and csv_col_for_placeholder != "Not Mapped" and csv_col_for_placeholder in row_data:
                        value_to_insert = row_data[csv_col_for_placeholder]
                    current_subject = current_subject.replace(placeholder, str(value_to_insert))
                    current_body = current_body.replace(placeholder, str(value_to_insert))
                emails_to_send_list.append({'recipient_email': recipient_email, 'subject': current_subject, 'body': current_body, 'row_identifier': f"CSV Row {i+1}"})
            if not emails_to_send_list: messagebox.showinfo("Info", "No valid recipient emails found in CSV data."); return

        schedule_dt_obj = None; date_to_use = self.ui_schedule_date_var.get(); time_to_use = self.ui_schedule_time_var.get()
        if not date_to_use and not time_to_use: 
            date_to_use = self.profile_schedule_date_var.get(); time_to_use = self.profile_schedule_time_var.get()
            if date_to_use or time_to_use: self.log_message(f"Using schedule from active profile: {date_to_use} {time_to_use}")

        if date_to_use or time_to_use:
            schedule_dt_obj = self._validate_schedule_datetime(date_to_use, time_to_use)
            if schedule_dt_obj == "Invalid": messagebox.showerror("Error", "Invalid Schedule Date/Time format."); return
            if schedule_dt_obj and schedule_dt_obj <= datetime.datetime.now(): messagebox.showerror("Schedule Error", f"Scheduled time is in the past."); return
        
        if scheduled_campaign_data:
             self._perform_email_sending(emails_to_send_list, is_test=False) 
             if campaign_id_to_send in self.scheduled_campaigns:
                del self.scheduled_campaigns[campaign_id_to_send]; self.save_scheduled_campaigns_to_file()
             return

        if schedule_dt_obj: 
            campaign_id = str(uuid.uuid4())
            self.scheduled_campaigns[campaign_id] = {
                "scheduled_datetime_str": schedule_dt_obj.strftime("%Y-%m-%d %H:%M:%S"),
                "emails_to_send_list": emails_to_send_list,
                # Store the settings at the time of scheduling
                "sender_email": sender_email_for_campaign, "sender_password": sender_password_for_campaign,
                "cv_path": cv_path_for_campaign, "status": "pending",
                "profile_name_at_schedule": self.active_profile_name.get()
            }
            self.save_scheduled_campaigns_to_file()
            self.log_message(f"Campaign ID {campaign_id} scheduled for {schedule_dt_obj.strftime('%Y-%m-%d %H:%M')}.")
            messagebox.showinfo("Scheduled", f"Email campaign scheduled for {schedule_dt_obj.strftime('%Y-%m-%d %H:%M')}.\nApp must be running for automatic sending.")
        else: 
            if not messagebox.askyesno("Confirm Send", f"Send {len(emails_to_send_list)} emails now?"): return
            self._perform_email_sending(emails_to_send_list, is_test=False)

    def send_manual_email_process(self):
        recipient_email = self.manual_email_var.get(); first_name = self.manual_first_name_var.get(); company_name = self.manual_company_name_var.get(); role = self.manual_role_var.get() 
        if not recipient_email or not self._is_valid_email(recipient_email): messagebox.showerror("Validation Error", "Valid recipient email is required."); return
        sender_email = self.smtp_email_var.get(); sender_password = self.smtp_password_var.get()
        if not sender_email or not sender_password: messagebox.showerror("Error", "Gmail address or App Password not set in active profile."); return
        if not self._is_valid_email(sender_email): messagebox.showerror("Error", "Invalid sender Gmail address format."); return
        subject_template = self.email_subject_var.get(); body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return
        cv_path = self.cv_file_path.get() 
        if cv_path and not os.path.exists(cv_path):
            if not messagebox.askyesno("CV Path Invalid", f"CV path invalid:\n'{cv_path}'\nContinue without CV?"): return
            cv_path = None 
        elif cv_path and not cv_path.lower().endswith(".pdf"): messagebox.showerror("Error", "CV file must be a PDF."); return
        elif not cv_path: self.log_message("No CV selected in active profile for manual send.", error=False)
        current_subject = subject_template.replace(DEFAULT_PLACEHOLDERS["FIRST_NAME"], first_name or "").replace(DEFAULT_PLACEHOLDERS["COMPANY_NAME"], company_name or "").replace(DEFAULT_PLACEHOLDERS["ROLE"], role or "").replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], "") 
        current_body = body_template.replace(DEFAULT_PLACEHOLDERS["FIRST_NAME"], first_name or "").replace(DEFAULT_PLACEHOLDERS["COMPANY_NAME"], company_name or "").replace(DEFAULT_PLACEHOLDERS["ROLE"], role or "").replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], "")
        email_details = [{'recipient_email': recipient_email, 'subject': current_subject, 'body': current_body, 'row_identifier': "Manual Send"}]
        if not messagebox.askyesno("Confirm Manual Send", f"Send email to {recipient_email} now?"): return
        self._perform_email_sending(email_details, is_test=True) 

    def send_test_email_process(self): 
        sender_email = self.smtp_email_var.get()
        if not sender_email or not self._is_valid_email(sender_email): messagebox.showerror("Error", "Your (sender) Gmail address is not set or invalid in active profile."); return
        subject_template = self.email_subject_var.get(); body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return
        test_fill_data = {}; active_tab_text = ""
        try: active_tab_text = self.notebook.tab(self.notebook.select(), "text")
        except tk.TclError: pass
        if active_tab_text == "Manual Send" and (self.manual_first_name_var.get() or self.manual_company_name_var.get()): 
            self.log_message("Preparing test email using data from 'Manual Send' tab inputs.")
            test_fill_data[DEFAULT_PLACEHOLDERS["FIRST_NAME"]] = self.manual_first_name_var.get() or "[TEST_FIRST_NAME]"; test_fill_data[DEFAULT_PLACEHOLDERS["LAST_NAME"]] = "" 
            test_fill_data[DEFAULT_PLACEHOLDERS["COMPANY_NAME"]] = self.manual_company_name_var.get() or "[TEST_COMPANY]"; test_fill_data[DEFAULT_PLACEHOLDERS["ROLE"]] = self.manual_role_var.get() or "[TEST_ROLE]"
        elif self.csv_data:
            self.log_message("Preparing test email using data from the first CSV row.")
            first_row = self.csv_data[0]
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                csv_col_name = self.column_mappings[key].get()
                if csv_col_name and csv_col_name != "Not Mapped" and csv_col_name in first_row: test_fill_data[placeholder] = first_row[csv_col_name]
                else: test_fill_data[placeholder] = f"[{key.upper()}_TEST_DATA]"
        else:
            self.log_message("Preparing test email using generic placeholder data (no CSV/Manual data).")
            for key, placeholder in DEFAULT_PLACEHOLDERS.items(): test_fill_data[placeholder] = f"[{key.upper()}_GENERIC_TEST]"
        current_subject = subject_template; current_body = body_template
        for placeholder, value in test_fill_data.items():
            current_subject = current_subject.replace(placeholder, str(value)); current_body = current_body.replace(placeholder, str(value))
        email_details = [{'recipient_email': sender_email, 'subject': f"[TEST EMAIL] {current_subject}", 'body': current_body, 'row_identifier': "Test Email"}]
        if not messagebox.askyesno("Confirm Test Send", f"Send a test email to yourself ({sender_email})?"): return
        self._perform_email_sending(email_details, is_test=True)

    # --- Custom Email Batch Methods ---
    def refresh_custom_emails_listbox(self):
        self.custom_emails_listbox.delete(0, tk.END)
        for i, email_data in enumerate(self.custom_email_batch):
            display_text = f"{i+1}. To: {email_data['recipient_email']} - Subject: {email_data['subject'][:30]}..."
            self.custom_emails_listbox.insert(tk.END, display_text)

    def add_or_edit_custom_email_dialog(self, edit_mode=False):
        index_to_edit = None
        existing_email_data = None
        if edit_mode:
            selected_indices = self.custom_emails_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("Info", "No email selected to edit.", parent=self.root); return
            index_to_edit = selected_indices[0]
            existing_email_data = self.custom_email_batch[index_to_edit]


        dialog = tk.Toplevel(self.root)
        dialog.title("Compose Custom Email" if not edit_mode else "Edit Custom Email")
        dialog.geometry("700x600")
        dialog.transient(self.root); dialog.grab_set()

        recipient_var = tk.StringVar(value=existing_email_data["recipient_email"] if existing_email_data else "")
        subject_var = tk.StringVar(value=existing_email_data["subject"] if existing_email_data else "")
        use_template_var = tk.BooleanVar(value=existing_email_data["use_template"] if existing_email_data else False)
        
        ph_vars = {key: tk.StringVar(value=existing_email_data["template_placeholders"].get(key, "") if existing_email_data and "template_placeholders" in existing_email_data and existing_email_data["use_template"] else "") for key in DEFAULT_PLACEHOLDERS if key != "LAST_NAME"} 

        ttk.Label(dialog, text="Recipient Email:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(dialog, textvariable=recipient_var, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        use_template_check = ttk.Checkbutton(dialog, text="Use Active Profile's Template", variable=use_template_var, command=lambda: toggle_template_fields())
        use_template_check.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        template_fields_frame = ttk.Frame(dialog)
        template_fields_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky="ew")

        ttk.Label(dialog, text="Subject:").grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        subject_entry = ttk.Entry(dialog, textvariable=subject_var, width=80)
        subject_entry.grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        
        ttk.Label(dialog, text="Body:").grid(row=4, column=0, padx=5, pady=5, sticky="nw")
        body_text_widget = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, height=15, width=80)
        body_text_widget.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky="nsew")
        if existing_email_data: body_text_widget.insert("1.0", existing_email_data["body"])

        dialog.columnconfigure(1, weight=1); dialog.rowconfigure(4, weight=1)

        def generate_from_template():
            if not use_template_var.get(): return
            tpl_subject = self.email_subject_var.get(); tpl_body = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
            for key, ph_tk_var in ph_vars.items():
                placeholder_text = DEFAULT_PLACEHOLDERS[key]
                tpl_subject = tpl_subject.replace(placeholder_text, ph_tk_var.get() or "")
                tpl_body = tpl_body.replace(placeholder_text, ph_tk_var.get() or "")
            tpl_subject = tpl_subject.replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], ""); tpl_body = tpl_body.replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], "")
            subject_var.set(tpl_subject); body_text_widget.delete("1.0", tk.END); body_text_widget.insert("1.0", tpl_body)

        def toggle_template_fields():
            for widget in template_fields_frame.winfo_children(): widget.destroy()
            if use_template_var.get():
                subject_entry.config(state="readonly"); body_text_widget.config(state="disabled")
                r = 0
                for key, ph_tk_var in ph_vars.items():
                    ttk.Label(template_fields_frame, text=f"{key.replace('_',' ').title()}:").grid(row=r, column=0, padx=2, pady=2, sticky="w")
                    ttk.Entry(template_fields_frame, textvariable=ph_tk_var, width=40).grid(row=r, column=1, padx=2, pady=2, sticky="ew"); r += 1
                ttk.Button(template_fields_frame, text="Generate from Template", command=generate_from_template).grid(row=r, column=0, columnspan=2, pady=5)
                generate_from_template()
            else:
                subject_entry.config(state="normal"); body_text_widget.config(state="normal")
                if not edit_mode or not existing_email_data or not existing_email_data["use_template"]:
                    if not subject_var.get() and not body_text_widget.get("1.0", tk.END).strip(): 
                        pass 
                    elif existing_email_data and existing_email_data["use_template"]: 
                         subject_var.set(""); body_text_widget.delete("1.0", tk.END)

        def save_custom_email():
            recipient = recipient_var.get(); subject = subject_var.get(); body = body_text_widget.get("1.0", tk.END).strip()
            if not recipient or not self._is_valid_email(recipient): messagebox.showerror("Error", "Valid recipient email is required.", parent=dialog); return
            if not subject.strip(): messagebox.showerror("Error", "Subject cannot be empty.", parent=dialog); return
            if not body.strip(): messagebox.showerror("Error", "Body cannot be empty.", parent=dialog); return
            email_entry = {
                "id": existing_email_data["id"] if existing_email_data else str(uuid.uuid4()),
                "recipient_email": recipient, "subject": subject, "body": body,
                "use_template": use_template_var.get(),
                "template_placeholders": {key: var.get() for key, var in ph_vars.items()} if use_template_var.get() else {}
            }
            if edit_mode and index_to_edit is not None: self.custom_email_batch[index_to_edit] = email_entry
            else: self.custom_email_batch.append(email_entry)
            self.refresh_custom_emails_listbox(); dialog.destroy()

        ttk.Button(dialog, text="Save Email to Batch", command=save_custom_email).grid(row=5, column=1, columnspan=2, pady=10)
        toggle_template_fields() 

    def remove_selected_custom_email(self):
        selected_indices = self.custom_emails_listbox.curselection()
        if not selected_indices: messagebox.showinfo("Info", "No email selected to remove.", parent=self.root); return
        if messagebox.askyesno("Confirm Remove", "Are you sure you want to remove the selected custom email from the batch?"):
            for index in sorted(selected_indices, reverse=True): del self.custom_email_batch[index]
            self.refresh_custom_emails_listbox()

    def clear_custom_email_batch(self):
        if not self.custom_email_batch: messagebox.showinfo("Info", "Custom email batch is already empty.", parent=self.root); return
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all emails from the custom batch?"):
            self.custom_email_batch.clear(); self.refresh_custom_emails_listbox()

    def send_custom_email_batch_process(self):
        if not self.custom_email_batch: messagebox.showinfo("Info", "No custom emails in the batch to send.", parent=self.root); return
        sender_email = self.smtp_email_var.get(); sender_password = self.smtp_password_var.get()
        if not sender_email or not sender_password: messagebox.showerror("Error", "Gmail address or App Password not set in active profile."); return
        if not self._is_valid_email(sender_email): messagebox.showerror("Error", "Invalid sender Gmail address format."); return
        cv_path = self.cv_file_path.get() 
        if cv_path and not os.path.exists(cv_path):
            if not messagebox.askyesno("CV Path Invalid", f"CV path invalid:\n'{cv_path}'\nContinue without CV?"): return
            cv_path = None 
        elif cv_path and not cv_path.lower().endswith(".pdf"): messagebox.showerror("Error", "CV file must be a PDF."); return
        elif not cv_path: self.log_message("No CV selected. Custom batch emails will be sent without attachments.", error=False)
        emails_to_send_list_for_smtp = []
        for i, custom_email_data in enumerate(self.custom_email_batch):
            emails_to_send_list_for_smtp.append({
                'recipient_email': custom_email_data['recipient_email'], 'subject': custom_email_data['subject'],
                'body': custom_email_data['body'], 'row_identifier': f"Custom Batch Email {i+1}"})
        if not messagebox.askyesno("Confirm Send Batch", f"Send {len(emails_to_send_list_for_smtp)} custom emails now?"): return
        self._perform_email_sending(emails_to_send_list_for_smtp, is_test=False, is_custom_batch=True)


    # --- Persistent Scheduling Methods ---
    def load_scheduled_campaigns_from_file(self):
        if os.path.exists(SCHEDULED_CAMPAIGNS_FILE):
            try:
                with open(SCHEDULED_CAMPAIGNS_FILE, "r") as f: return json.load(f)
            except (IOError, json.JSONDecodeError):
                self.log_message(f"Error loading scheduled campaigns file or file corrupted. Starting with empty schedule.", error=True); return {}
        return {}

    def save_scheduled_campaigns_to_file(self):
        try:
            with open(SCHEDULED_CAMPAIGNS_FILE, "w") as f: json.dump(self.scheduled_campaigns, f, indent=4)
        except IOError as e: self.log_message(f"Error saving scheduled campaigns data: {e}", error=True)

    def periodic_schedule_check(self):
        self.check_for_pending_scheduled_jobs(silent=True)
        self.root.after(60000, self.periodic_schedule_check) 

    def check_for_pending_scheduled_jobs(self, silent=False):
        now = datetime.datetime.now(); campaigns_to_send_now = {}
        for campaign_id, campaign_data in list(self.scheduled_campaigns.items()):
            if campaign_data.get("status") == "pending":
                try:
                    scheduled_dt = datetime.datetime.strptime(campaign_data["scheduled_datetime_str"], "%Y-%m-%d %H:%M:%S")
                    if now >= scheduled_dt: campaigns_to_send_now[campaign_id] = campaign_data
                except ValueError:
                    self.log_message(f"Campaign {campaign_id} has invalid datetime string. Skipping.", error=True)
                    campaign_data["status"] = "error_invalid_date"
        if campaigns_to_send_now:
            if not silent:
                ids_str = ", ".join(campaigns_to_send_now.keys())
                if messagebox.askyesno("Pending Scheduled Emails", f"Scheduled campaign(s) due: {ids_str}.\nSend now?"):
                    for cid, cdata in campaigns_to_send_now.items():
                        self.log_message(f"Executing overdue scheduled campaign: {cid}")
                        self.send_emails_process(campaign_id_to_send=cid, scheduled_campaign_data=cdata)
                else: self.log_message(f"User chose not to send overdue campaigns: {ids_str}")
            else: 
                 for cid, cdata in campaigns_to_send_now.items():
                    self.log_message(f"Silently executing overdue scheduled campaign: {cid}")
                    self.send_emails_process(campaign_id_to_send=cid, scheduled_campaign_data=cdata)
        elif not silent: self.log_message("No overdue scheduled email campaigns found.")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        try:
            style = ttk.Style(root); available_themes = style.theme_names()
            if 'clam' in available_themes: style.theme_use('clam')
            elif 'vista' in available_themes and os.name == 'nt': style.theme_use('vista')
            elif 'aqua' in available_themes and os.name == 'posix': style.theme_use('aqua') 
        except Exception as e: print(f"Could not set theme: {e}")
        app = BulkEmailerApp(root); root.mainloop()
    except tk.TclError as e: print(f"Tkinter TclError: {e}\nThis script requires a graphical display environment.")
    except Exception as e: print(f"An unexpected error occurred: {e}"); import traceback; traceback.print_exc()

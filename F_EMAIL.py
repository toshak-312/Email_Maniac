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

CONFIG_FILE = "bulk_emailer_config_profiles.json"
DEFAULT_PLACEHOLDERS = {
    "FIRST_NAME": "{{FIRST_NAME}}",
    "LAST_NAME": "{{LAST_NAME}}",
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
        self.root.title("Advanced Bulk Emailer (Profiles)")
        self.root.geometry("950x750")

        # --- Profile Management ---
        self.profiles = {}
        self.active_profile_name = tk.StringVar()
        self.profile_keys_for_dropdown = []

        # --- Data (tied to active profile or session) ---
        self.csv_file_paths = [] # List of paths for current session
        self.cv_file_path = tk.StringVar() # From active profile
        self.csv_headers = [] # Combined from all loaded CSVs
        self.csv_data = [] # Combined from all loaded CSVs

        # --- Column Mapping (tied to active profile) ---
        self.email_column_var = tk.StringVar()
        self.column_mappings = {key: tk.StringVar() for key in DEFAULT_PLACEHOLDERS}

        # --- Email Content (tied to active profile) ---
        self.email_subject_var = tk.StringVar()
        self.email_body_text_widget = None

        # --- SMTP Settings (tied to active profile) ---
        self.smtp_email_var = tk.StringVar()
        self.smtp_password_var = tk.StringVar()

        # --- Scheduling (tied to active profile) ---
        self.preferred_send_time_var = tk.StringVar() # HH:MM format

        # --- Manual Send ---
        self.manual_email_var = tk.StringVar()
        self.manual_first_name_var = tk.StringVar()
        self.manual_company_name_var = tk.StringVar()
        self.manual_role_var = tk.StringVar() # Added for consistency

        self.load_app_config() # Load profiles and active profile name
        self.create_widgets()
        
        if not self.profiles: # Ensure at least a default profile exists
            self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True)
        elif self.active_profile_name.get() not in self.profiles:
            # If saved active profile doesn't exist, pick first available or create default
            if self.profile_keys_for_dropdown:
                self.active_profile_name.set(self.profile_keys_for_dropdown[0])
            else: # Should not happen if default is always created
                 self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True)
        
        self.load_profile_data(self.active_profile_name.get())
        self.update_column_mapping_dropdowns_state()


    def get_default_profile_settings(self):
        """Returns a dictionary with truly default settings for a new profile (blank SMTP)."""
        return {
            "csv_file_paths": [], 
            "cv_file_path": "",
            "email_column": "",
            "column_mappings": {key: "" for key in DEFAULT_PLACEHOLDERS},
            "email_subject": "Internship Application: {{ROLE}} at {{COMPANY_NAME}}",
            "email_body": ("Dear Hiring Manager at {{COMPANY_NAME}},\n\n"
                           "I am writing to express my keen interest in an internship opportunity, potentially in a {{ROLE}} capacity or a related field, at {{COMPANY_NAME}}.\n\n"
                           "My name is {{FIRST_NAME}} {{LAST_NAME}}, and I am a highly motivated student with a passion for [Your Field/Area of Interest].\n\n"
                           "I have attached my CV for your review.\n\n"
                           "Thank you for your time and consideration.\n\n"
                           "Sincerely,\n"
                           "{{FIRST_NAME}} {{LAST_NAME}}"),
            "smtp_email": "", # Starts blank, will be inherited if possible
            "smtp_password": "", # Starts blank
            "preferred_send_time": "" 
        }

    def save_app_config(self):
        """Saves all profiles and the active profile name."""
        # Ensure the currently active UI data is saved to the active profile object
        self.save_current_profile_data_to_object()

        app_config = {
            "active_profile_name": self.active_profile_name.get(),
            "profiles": self.profiles
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(app_config, f, indent=4)
            self.log_message("Application configuration (all profiles) saved.")
        except IOError as e:
            self.log_message(f"Error saving application configuration: {e}", error=True)
        except Exception as e: # Catch any other unexpected error during save
            self.log_message(f"Unexpected error during save_app_config: {e}", error=True)


    def load_app_config(self):
        """Loads all profiles and the last active profile name."""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f:
                    app_config = json.load(f)
                self.active_profile_name.set(app_config.get("active_profile_name", DEFAULT_PROFILE_NAME))
                self.profiles = app_config.get("profiles", {})
                
                if not self.profiles: # If profiles dict is empty after loading
                    self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                    if not self.active_profile_name.get(): # If active_profile_name was also empty
                        self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                
                self.profile_keys_for_dropdown = list(self.profiles.keys())
                if not self.profile_keys_for_dropdown: # Should not happen if default is created
                     self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                     self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                     self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]


            else: # No config file, create a default profile
                self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
                self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
                # self.save_app_config() # Save the initial default config - deferred to first save action or quit

        except (IOError, json.JSONDecodeError) as e:
            self.log_message(f"Error loading config or config corrupted: {e}. Creating default.", error=True)
            self.active_profile_name.set(DEFAULT_PROFILE_NAME)
            self.profiles = {DEFAULT_PROFILE_NAME: self.get_default_profile_settings()}
            self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]


    def save_current_profile_data_to_object(self):
        """Updates the self.profiles dictionary with data from the UI for the active profile."""
        profile_name = self.active_profile_name.get()
        if not profile_name or profile_name not in self.profiles:
            # This case should ideally not happen if a profile is always active
            # self.log_message("No active profile selected or profile not found to save data.", error=True)
            return

        # Ensure the profile entry exists, if not, create it (should be defensive)
        if profile_name not in self.profiles:
            self.profiles[profile_name] = self.get_default_profile_settings() # Initialize with defaults
            # Potentially inherit SMTP from a previously active one if logic allows

        current_profile_data = self.profiles[profile_name] # Get a reference
        
        current_profile_data["cv_file_path"] = self.cv_file_path.get()
        current_profile_data["email_column"] = self.email_column_var.get()
        current_profile_data["column_mappings"] = {key: var.get() for key, var in self.column_mappings.items()}
        current_profile_data["email_subject"] = self.email_subject_var.get()
        if self.email_body_text_widget: # Check if widget exists
            current_profile_data["email_body"] = self.email_body_text_widget.get("1.0", tk.END).strip()
        else: # Fallback if widget not ready (e.g. early call)
            current_profile_data["email_body"] = self.profiles[profile_name].get("email_body","")


        current_profile_data["smtp_email"] = self.smtp_email_var.get()
        current_profile_data["smtp_password"] = self.smtp_password_var.get()
        current_profile_data["preferred_send_time"] = self.preferred_send_time_var.get()
        
        # self.profiles[profile_name] is already the reference, so changes are direct.
        # self.log_message(f"Profile '{profile_name}' data updated in memory.")


    def load_profile_data(self, profile_name):
        """Loads data from the specified profile into the UI elements."""
        if not profile_name or profile_name not in self.profiles:
            self.log_message(f"Profile '{profile_name}' not found. Cannot load.", error=True)
            # Attempt to load default profile if current one is invalid
            if DEFAULT_PROFILE_NAME in self.profiles and profile_name != DEFAULT_PROFILE_NAME:
                self.active_profile_name.set(DEFAULT_PROFILE_NAME)
                self.load_profile_data(DEFAULT_PROFILE_NAME)
            return

        profile_data = self.profiles[profile_name]
        self.active_profile_name.set(profile_name) 
        
        self.cv_file_path.set(profile_data.get("cv_file_path", ""))
        self.email_column_var.set(profile_data.get("email_column", ""))
        
        loaded_mappings = profile_data.get("column_mappings", {})
        for key, var_tk in self.column_mappings.items():
            var_tk.set(loaded_mappings.get(key, "")) # Default to "" if key missing in saved data
            
        self.email_subject_var.set(profile_data.get("email_subject", "Internship Application: {{ROLE}} at {{COMPANY_NAME}}"))
        if self.email_body_text_widget: # Ensure widget exists before trying to set its content
            self.email_body_text_widget.delete("1.0", tk.END)
            self.email_body_text_widget.insert("1.0", profile_data.get("email_body", self.get_default_profile_settings()["email_body"]))
        
        self.smtp_email_var.set(profile_data.get("smtp_email", ""))
        self.smtp_password_var.set(profile_data.get("smtp_password", ""))
        self.preferred_send_time_var.set(profile_data.get("preferred_send_time", ""))

        self.update_column_mapping_dropdowns() # Update dropdowns based on current CSV headers (if any) and loaded mappings
        self.log_message(f"Profile '{profile_name}' loaded.")


    def on_profile_selected(self, event=None):
        """Called when a profile is selected from the dropdown."""
        # The save of the *previous* profile's UI state should happen before loading the new one.
        # This is tricky. The `save_current_profile_data_to_object` is called by `save_app_config`.
        # For explicit profile switch, we should save the outgoing profile's UI state.
        # However, OptionMenu command doesn't easily allow knowing the *previous* value.
        # Simplest for now: rely on the "Save Current Profile" button or auto-save on quit.
        # Or, when switching, we could iterate all profiles and save their current UI counterparts,
        # but that's complex if UI elements are shared/reconfigured.

        # The current approach: load_profile_data populates UI from the selected profile.
        # Changes made in UI are only saved to the profile object on "Save" or "Quit".
        selected_profile = self.active_profile_name.get()
        self.load_profile_data(selected_profile)

    def create_new_profile_dialog(self):
        profile_name = simpledialog.askstring("New Profile", "Enter name for the new profile:", parent=self.root)
        if profile_name:
            if profile_name in self.profiles:
                messagebox.showerror("Error", f"Profile '{profile_name}' already exists.")
            else:
                self.create_new_profile(profile_name, make_active=True)
    
    def create_new_profile(self, profile_name, make_active=False, initial_setup=False):
        """Creates a new profile. If not initial_setup, inherits SMTP from current active."""
        new_profile_settings = self.get_default_profile_settings() # Base defaults

        if not initial_setup: # Inherit SMTP if not the very first profile creation
            current_active_profile_name_for_inheritance = self.active_profile_name.get()
            if current_active_profile_name_for_inheritance and current_active_profile_name_for_inheritance in self.profiles:
                active_profile_data = self.profiles[current_active_profile_name_for_inheritance]
                new_profile_settings["smtp_email"] = active_profile_data.get("smtp_email", "")
                new_profile_settings["smtp_password"] = active_profile_data.get("smtp_password", "")
                self.log_message(f"New profile '{profile_name}' inherited SMTP settings from '{current_active_profile_name_for_inheritance}'.")

        self.profiles[profile_name] = new_profile_settings
        self.profile_keys_for_dropdown = list(self.profiles.keys())
        self.update_profile_dropdown()

        if make_active:
            self.active_profile_name.set(profile_name)
            self.load_profile_data(profile_name) # Load the new profile's data into UI

        self.log_message(f"Profile '{profile_name}' created.")
        if not initial_setup: # Don't save during initial app load sequence
             self.save_app_config() 


    def delete_current_profile_dialog(self):
        profile_name_to_delete = self.active_profile_name.get()
        if not profile_name_to_delete or profile_name_to_delete == DEFAULT_PROFILE_NAME:
            messagebox.showerror("Error", "Cannot delete the default profile or no profile selected.")
            return
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete profile '{profile_name_to_delete}'? This cannot be undone."):
            if profile_name_to_delete in self.profiles:
                del self.profiles[profile_name_to_delete]
                self.profile_keys_for_dropdown = list(self.profiles.keys())
                
                new_active = DEFAULT_PROFILE_NAME if DEFAULT_PROFILE_NAME in self.profiles else (self.profile_keys_for_dropdown[0] if self.profile_keys_for_dropdown else "")
                self.active_profile_name.set(new_active)
                self.update_profile_dropdown() # Update dropdown choices
                
                if new_active:
                    self.load_profile_data(new_active) # Load the new active profile
                else: # This case means all profiles were deleted, including default (which shouldn't happen)
                    self.create_new_profile(DEFAULT_PROFILE_NAME, make_active=True, initial_setup=True) # Recreate default
                
                self.log_message(f"Profile '{profile_name_to_delete}' deleted.")
                self.save_app_config()

    def update_profile_dropdown(self):
        menu = self.profile_menu['menu']
        menu.delete(0, 'end')
        
        # Ensure there's always at least one profile key (e.g., Default Profile)
        if not self.profile_keys_for_dropdown and DEFAULT_PROFILE_NAME not in self.profiles:
            # This is a fallback, should be handled by load_app_config or create_new_profile
            self.profiles[DEFAULT_PROFILE_NAME] = self.get_default_profile_settings()
            self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]
            if not self.active_profile_name.get():
                self.active_profile_name.set(DEFAULT_PROFILE_NAME)

        elif not self.profile_keys_for_dropdown and DEFAULT_PROFILE_NAME in self.profiles:
             self.profile_keys_for_dropdown = [DEFAULT_PROFILE_NAME]


        for profile_key in self.profile_keys_for_dropdown:
            menu.add_command(label=profile_key, command=lambda pk=profile_key: self.set_and_load_profile(pk))
        
        current_active = self.active_profile_name.get()
        if current_active not in self.profile_keys_for_dropdown:
             if self.profile_keys_for_dropdown:
                self.active_profile_name.set(self.profile_keys_for_dropdown[0])
             # else: self.active_profile_name.set("") # No profiles exist
        # self.active_profile_name.set(self.active_profile_name.get()) # Refresh the display of OptionMenu

    def set_and_load_profile(self, profile_key):
        """Helper to set active profile and then load its data, used by OptionMenu."""
        self.active_profile_name.set(profile_key)
        self.on_profile_selected()


    def create_widgets(self):
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tab_profile_csv = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_profile_csv, text='Profiles & CSV')
        self.create_tab_profile_csv(tab_profile_csv)

        tab_mapping = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_mapping, text='Column Mapping')
        self.create_tab_mapping(tab_mapping)

        tab_email_content = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_email_content, text='Email Content')
        self.create_tab_email_content(tab_email_content)
        
        tab_manual_send = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_manual_send, text='Manual Send')
        self.create_tab_manual_send(tab_manual_send)

        tab_settings_send = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_settings_send, text='Settings & Send')
        self.create_tab_settings_send(tab_settings_send)
        
        log_frame = ttk.LabelFrame(main_container, text="Log", padding=10)
        log_frame.pack(fill=tk.X, padx=5, pady=(10,5), side=tk.BOTTOM)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=5, state='disabled', font=("Helvetica", 9))
        self.log_text.pack(fill=tk.X, expand=False)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_profile_dropdown() # Initialize profile dropdown content after widgets created


    def create_tab_profile_csv(self, parent_tab):
        # Profile Management Frame
        profile_frame = ttk.LabelFrame(parent_tab, text="User Profiles", padding=10)
        profile_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        ttk.Label(profile_frame, text="Active Profile:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Ensure active_profile_name has a valid initial value for OptionMenu
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

        # File Selection Frame
        file_frame = ttk.LabelFrame(parent_tab, text="Load Data & CV (for current session/profile)", padding=10)
        file_frame.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        
        ttk.Button(file_frame, text="Load CSV File(s)", command=self.load_csv_files).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.csv_paths_label = ttk.Label(file_frame, text="No CSVs loaded.", wraplength=350)
        self.csv_paths_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Button(file_frame, text="Select CV (PDF for active profile)", command=self.select_cv_file).grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(file_frame, textvariable=self.cv_file_path, wraplength=350).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        file_frame.columnconfigure(1, weight=1)
        
        parent_tab.columnconfigure(0, weight=1)

    def create_tab_mapping(self, parent_tab):
        mapping_frame = ttk.LabelFrame(parent_tab, text="Map CSV Columns (for active profile)", padding=10)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        ttk.Label(mapping_frame, text="CSV Column for Email Address (Required):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.email_column_menu = ttk.OptionMenu(mapping_frame, self.email_column_var, self.email_column_var.get() or "Select Email Column", *(self.csv_headers or ["Not Mapped"]))
        self.email_column_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(mapping_frame, text="--- Map Placeholders to CSV Columns (Auto-detected where possible): ---").grid(row=1, column=0, columnspan=2, pady=(10,2))
        
        self.placeholder_menus = {}
        current_row = 2
        for key, placeholder_text in DEFAULT_PLACEHOLDERS.items():
            label_text = f"{key.replace('_', ' ').title()} ({placeholder_text}):"
            ttk.Label(mapping_frame, text=label_text).grid(row=current_row, column=0, padx=5, pady=3, sticky="w")
            var = self.column_mappings[key]
            # Ensure var has a value for OptionMenu initialization
            initial_val_map = var.get() if var.get() else ("Not Mapped" if not self.csv_headers else (self.csv_headers[0] if self.csv_headers else "Not Mapped"))
            menu = ttk.OptionMenu(mapping_frame, var, initial_val_map, *(self.csv_headers if self.csv_headers else ["Not Mapped"]))
            menu.grid(row=current_row, column=1, padx=5, pady=3, sticky="ew")
            self.placeholder_menus[key] = menu
            current_row += 1
        
        mapping_frame.columnconfigure(1, weight=1)

    def create_tab_email_content(self, parent_tab):
        email_template_frame = ttk.LabelFrame(parent_tab, text="Email Template Editor (for active profile)", padding=10)
        email_template_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(email_template_frame, text="Subject:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(email_template_frame, textvariable=self.email_subject_var, width=80).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(email_template_frame, text="Body:").grid(row=1, column=0, padx=5, pady=5, sticky="nw")
        self.email_body_text_widget = scrolledtext.ScrolledText(email_template_frame, wrap=tk.WORD, height=15, width=80, font=("Helvetica", 10))
        self.email_body_text_widget.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        # Initial body text is loaded via load_profile_data
        
        ttk.Button(email_template_frame, text="Preview Email (using first CSV row if available)", command=self.preview_email).grid(row=2, column=1, padx=5, pady=10, sticky="e")
        
        email_template_frame.columnconfigure(1, weight=1)
        email_template_frame.rowconfigure(1, weight=1)

    def create_tab_manual_send(self, parent_tab):
        manual_frame = ttk.LabelFrame(parent_tab, text="Send Single Email Manually", padding=10)
        manual_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(manual_frame, text="Recipient Email:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(manual_frame, textvariable=self.manual_email_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(manual_frame, text="{{FIRST_NAME}}:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(manual_frame, textvariable=self.manual_first_name_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(manual_frame, text="{{COMPANY_NAME}}:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(manual_frame, textvariable=self.manual_company_name_var, width=50).grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(manual_frame, text="{{ROLE}} (Optional):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(manual_frame, textvariable=self.manual_role_var, width=50).grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        manual_frame.columnconfigure(1, weight=1)

        action_buttons_frame = ttk.Frame(manual_frame)
        action_buttons_frame.grid(row=4, column=0, columnspan=2, pady=15)

        ttk.Button(action_buttons_frame, text="Preview Manual Email", command=lambda: self.preview_email(manual_mode=True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_buttons_frame, text="Send Manual Email", command=self.send_manual_email_process, style="Accent.TButton").pack(side=tk.LEFT, padx=5)


    def create_tab_settings_send(self, parent_tab):
        # SMTP Settings Frame
        smtp_frame = ttk.LabelFrame(parent_tab, text="Gmail SMTP Settings (for active profile)", padding=10)
        smtp_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        ttk.Label(smtp_frame, text="Your Gmail Address:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(smtp_frame, textvariable=self.smtp_email_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(smtp_frame, text="Gmail App Password:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(smtp_frame, textvariable=self.smtp_password_var, show="*", width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        smtp_frame.columnconfigure(1, weight=1)

        # Scheduling Frame
        schedule_frame = ttk.LabelFrame(parent_tab, text="Scheduling (for active profile)", padding=10)
        schedule_frame.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        ttk.Label(schedule_frame, text="Preferred Send Time (HH:MM, 24-hr format, blank for immediate):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(schedule_frame, textvariable=self.preferred_send_time_var, width=10).grid(row=0, column=1, padx=5, pady=5, sticky="w")


        # Action Frame
        action_frame = ttk.LabelFrame(parent_tab, text="Bulk Sending Actions", padding=10)
        action_frame.grid(row=0, column=1, padx=15, pady=5, sticky="ns", rowspan=3) # rowspan to align better if needed

        self.send_button = ttk.Button(action_frame, text="Send Bulk Emails", command=self.send_emails_process, style="Accent.TButton")
        self.send_button.pack(pady=10, fill=tk.X, ipady=4)
        ttk.Button(action_frame, text="Send Test Email to Myself", command=self.send_test_email_process).pack(pady=5, fill=tk.X)
        
        # Progress Bar
        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.pack(pady=10, fill=tk.X)
        
        parent_tab.columnconfigure(0, weight=1) # SMTP and Schedule frames can expand
        parent_tab.columnconfigure(1, weight=0) # Action frame fixed width relative to its content

        style = ttk.Style()
        style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'))
        try: # Attempt to style button, may be theme-dependent
            style.map("Accent.TButton", foreground=[('!disabled', 'white')], background=[('active', 'darkgreen'), ('!disabled', 'green')])
        except tk.TclError:
            self.log_message("Note: Theme may not fully support custom button styling.", error=False)

    def on_closing(self):
        """Handles window close event by saving all settings automatically."""
        self.log_message("Closing application. Auto-saving all profiles and settings...")
        self.save_app_config() # Save everything
        self.root.destroy()

    def log_message(self, message, error=False):
        if not hasattr(self, 'log_text') or self.log_text is None:
            print(f"LOG ({'ERROR' if error else 'INFO'}): {message}")
            return
        try:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} - {message}\n", "error_tag" if error else "info_tag")
            self.log_text.tag_config("error_tag", foreground="red")
            self.log_text.tag_config("info_tag", foreground="black") # Or another non-red color
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
            if self.root and self.root.winfo_exists(): # Check if root window still exists
                 self.root.update_idletasks()
        except tk.TclError: # Handle cases where widget might be destroyed during shutdown
            print(f"LOG (TclError suppressed): {message}")


    def _auto_detect_columns(self):
        if not self.csv_headers: return

        # Email Column (prioritize profile setting if exists and valid)
        current_email_col_setting = self.email_column_var.get()
        if not current_email_col_setting or current_email_col_setting not in self.csv_headers:
            detected = False
            for header in self.csv_headers:
                if header.lower().replace(" ", "").replace("_", "") in AUTO_DETECT_PATTERNS["email_column"]:
                    self.email_column_var.set(header)
                    self.log_message(f"Auto-detected Email column: '{header}'")
                    detected = True
                    break
            if not detected and self.csv_headers : # If nothing specific found, and we have headers, pick first as placeholder
                 pass # Let user pick, or it will use current value which might be "Not Mapped"
        
        for key, patterns in AUTO_DETECT_PATTERNS.items():
            if key == "email_column": continue # Already handled
            
            current_mapping = self.column_mappings[key].get()
            # Only auto-detect if not set from profile or if current setting is not a valid header
            if not current_mapping or current_mapping == "Not Mapped" or current_mapping not in self.csv_headers:
                detected_placeholder = False
                for header in self.csv_headers:
                    normalized_header = header.lower().replace(" ", "").replace("_", "")
                    if normalized_header in patterns:
                        self.column_mappings[key].set(header)
                        self.log_message(f"Auto-detected {key.replace('_',' ').title()} column: '{header}'")
                        detected_placeholder = True
                        break
                if not detected_placeholder: # If not detected, ensure it's "Not Mapped" if no valid mapping exists
                    if self.column_mappings[key].get() not in self.csv_headers:
                         self.column_mappings[key].set("Not Mapped")


        self.update_column_mapping_dropdowns() # Refresh dropdowns with new auto-detections


    def _load_csv_data_from_paths(self, file_paths, silent=False):
        self.csv_data = []
        combined_headers = set()
        all_rows = []

        if not file_paths:
            self.csv_headers = []
            self.csv_paths_label.config(text="No CSVs loaded.")
            self.update_column_mapping_dropdowns()
            return True 

        for file_path in file_paths:
            try:
                with open(file_path, mode='r', encoding='utf-8-sig', newline='') as file:
                    reader = csv.DictReader(file)
                    if not reader.fieldnames:
                        if not silent: messagebox.showwarning("CSV Warning", f"CSV file '{os.path.basename(file_path)}' is empty or has no headers. Skipping.")
                        continue
                    
                    current_file_rows = list(reader) # Read all rows to check count
                    if not current_file_rows and not silent:
                         messagebox.showwarning("CSV Warning", f"CSV file '{os.path.basename(file_path)}' has headers but no data rows.")
                    
                    all_rows.extend(current_file_rows) # Add them even if empty, header processing is separate
                    for header in reader.fieldnames: # Add all headers from this file
                        combined_headers.add(header)
                if not silent: self.log_message(f"Successfully processed {os.path.basename(file_path)}.")
            except Exception as e:
                if not silent:
                    messagebox.showerror("CSV Error", f"Failed to load/parse {os.path.basename(file_path)}: {e}")
                    self.log_message(f"Failed to load {os.path.basename(file_path)}: {e}", error=True)
        
        self.csv_headers = sorted(list(combined_headers)) # Use unique sorted headers
        self.csv_data = all_rows # Assign combined rows

        if not self.csv_data and not silent and file_paths:
             self.log_message("Warning: All loaded CSVs combined resulted in no data rows (only headers might be present).", error=False) # Not a critical error
        elif self.csv_data:
             self.log_message(f"Total {len(self.csv_data)} rows loaded from {len(file_paths)} CSV file(s).")
        
        self.csv_paths_label.config(text=f"{len(file_paths)} CSV(s) loaded: " + ", ".join([os.path.basename(p) for p in file_paths]) if file_paths else "No CSVs loaded.")
        self._auto_detect_columns() # This will also call update_column_mapping_dropdowns
        return True


    def load_csv_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Select CSV Files",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if filepaths:
            self.csv_file_paths = list(filepaths) 
            self._load_csv_data_from_paths(self.csv_file_paths)


    def select_cv_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CV (PDF File for Active Profile)",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if file_path:
            if file_path.lower().endswith(".pdf"):
                self.cv_file_path.set(file_path) 
                self.log_message(f"CV selected for current profile: {os.path.basename(file_path)}")
                # Changes to cv_file_path are saved with profile on app_config save.
            else:
                messagebox.showerror("File Error", "Please select a PDF file for the CV.")

    def update_column_mapping_dropdowns(self):
        """Updates OptionMenu widgets with current CSV headers and preserves selections if possible."""
        options = ["Not Mapped"] + (self.csv_headers if self.csv_headers else [])
        
        # Update Email Column Dropdown
        if hasattr(self, 'email_column_menu'):
            current_email_col_val = self.email_column_var.get()
            self.email_column_menu['menu'].delete(0, 'end')
            # Set default for OptionMenu if current value is not in new options
            default_email_option = current_email_col_val if current_email_col_val in options else options[0]
            self.email_column_var.set(default_email_option) # Set the var first
            for option_val in options: # Then populate menu
                self.email_column_menu['menu'].add_command(label=option_val, command=tk._setit(self.email_column_var, option_val))
        
        # Update Placeholder Mapping Dropdowns
        if hasattr(self, 'placeholder_menus'):
            for key, menu_widget in self.placeholder_menus.items():
                current_placeholder_val = self.column_mappings[key].get()
                menu_widget['menu'].delete(0, 'end')
                default_placeholder_option = current_placeholder_val if current_placeholder_val in options else options[0]
                self.column_mappings[key].set(default_placeholder_option) # Set var first
                for option_val in options:
                    menu_widget['menu'].add_command(label=option_val, command=tk._setit(self.column_mappings[key], option_val))
        
        self.update_column_mapping_dropdowns_state()


    def update_column_mapping_dropdowns_state(self):
        state = tk.NORMAL if self.csv_headers else tk.DISABLED
        if hasattr(self, 'email_column_menu'): self.email_column_menu.config(state=state)
        if hasattr(self, 'placeholder_menus'):
            for menu in self.placeholder_menus.values(): menu.config(state=state)

    def _is_valid_email(self, email_string):
        if not email_string or not isinstance(email_string, str): return False
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        return re.match(pattern, email_string) is not None

    def _validate_send_time_format(self, time_str):
        if not time_str: return True 
        try:
            datetime.datetime.strptime(time_str, "%H:%M")
            return True
        except ValueError:
            return False

    def preview_email(self, manual_mode=False):
        if self.email_body_text_widget is None: messagebox.showerror("Error", "Email body editor not available."); return

        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END)
        
        preview_fill_data = {}

        if manual_mode:
            # For manual mode preview, use the manual input fields
            # No need to check manual_email_var.get() here, as it's just for preview
            preview_fill_data[DEFAULT_PLACEHOLDERS["FIRST_NAME"]] = self.manual_first_name_var.get() or "[MANUAL_FIRST_NAME]"
            preview_fill_data[DEFAULT_PLACEHOLDERS["LAST_NAME"]] = "" 
            preview_fill_data[DEFAULT_PLACEHOLDERS["COMPANY_NAME"]] = self.manual_company_name_var.get() or "[MANUAL_COMPANY_NAME]"
            preview_fill_data[DEFAULT_PLACEHOLDERS["ROLE"]] = self.manual_role_var.get() or "[MANUAL_ROLE]"
        else: # Bulk mode preview
            if not self.csv_data: messagebox.showinfo("Preview Info", "Load CSV data to preview bulk email."); return
            first_row = self.csv_data[0]
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                csv_col_name = self.column_mappings[key].get()
                if csv_col_name and csv_col_name != "Not Mapped" and csv_col_name in first_row:
                    preview_fill_data[placeholder] = first_row[csv_col_name]
                else: # If not mapped or data missing in first row for that column
                    preview_fill_data[placeholder] = f"[{key.upper()}_DATA]"
        
        final_subject = subject_template
        final_body = body_template
        for placeholder, value in preview_fill_data.items():
            final_subject = final_subject.replace(placeholder, str(value))
            final_body = final_body.replace(placeholder, str(value))
        
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Email Preview")
        preview_window.geometry("600x450")
        preview_window.transient(self.root); preview_window.grab_set()

        ttk.Label(preview_window, text="Subject:", font=("Helvetica", 11, "bold")).pack(pady=(10,2), anchor="w", padx=10)
        ttk.Label(preview_window, text=final_subject, wraplength=580, font=("Helvetica", 10)).pack(pady=(0,10), anchor="w", padx=10)
        ttk.Separator(preview_window, orient='horizontal').pack(fill='x', padx=10, pady=5)
        ttk.Label(preview_window, text="Body:", font=("Helvetica", 11, "bold")).pack(pady=(5,2), anchor="w", padx=10)
        body_prev_text = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, height=15, relief=tk.SOLID, borderwidth=1, font=("Helvetica", 10))
        body_prev_text.insert(tk.END, final_body)
        body_prev_text.config(state='disabled')
        body_prev_text.pack(pady=(0,10), padx=10, fill=tk.BOTH, expand=True)
        ttk.Button(preview_window, text="Close", command=preview_window.destroy).pack(pady=10)


    def _perform_email_sending(self, emails_to_send_list, is_test=False):
        cv_path = self.cv_file_path.get() 
        sender_email = self.smtp_email_var.get()
        sender_password = self.smtp_password_var.get()

        # CV path check for bulk sending (already done before calling for bulk, but good to have here too)
        if not is_test and cv_path and not os.path.exists(cv_path):
             self.log_message(f"CV path '{cv_path}' is invalid. Sending without CV.", error=True)
             cv_path = None # Ensure CV is not attached if path invalid during actual send
        elif not is_test and cv_path and not cv_path.lower().endswith(".pdf"):
             self.log_message(f"CV file '{cv_path}' is not a PDF. Sending without CV.", error=True)
             cv_path = None


        self.log_message(f"Starting SMTP process for {len(emails_to_send_list)} email(s)...")
        if hasattr(self, 'send_button'): self.send_button.config(state=tk.DISABLED)
        if hasattr(self, 'progress_bar'): self.progress_bar['value'] = 0; self.progress_bar['maximum'] = len(emails_to_send_list) if emails_to_send_list else 1

        sent_count = 0
        failed_count = 0

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.ehlo() # Greet server
            server.starttls()
            server.ehlo() # Greet again after TLS
            server.login(sender_email, sender_password)
            self.log_message("Logged into Gmail SMTP server.")

            for i, email_details in enumerate(emails_to_send_list):
                recipient_email = email_details['recipient_email']
                current_subject = email_details['subject']
                current_body = email_details['body']
                row_identifier = email_details.get('row_identifier', f"item {i+1}")

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg['Subject'] = current_subject
                msg.attach(MIMEText(current_body, 'plain', 'utf-8'))

                if cv_path and os.path.exists(cv_path) and cv_path.lower().endswith(".pdf"): 
                    try:
                        with open(cv_path, "rb") as attachment_file:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment_file.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(cv_path)}")
                        msg.attach(part)
                    except Exception as e:
                        self.log_message(f"Failed to attach CV for {recipient_email} ({row_identifier}): {e}", error=True)
                        if not is_test: failed_count += 1; self.update_progress(i + 1); continue 
                
                try:
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                    self.log_message(f"Email sent to {recipient_email} ({row_identifier})")
                    sent_count += 1
                except Exception as e:
                    self.log_message(f"Failed to send email to {recipient_email} ({row_identifier}): {e}", error=True)
                    if not is_test: failed_count += 1
                
                if not is_test: self.update_progress(i + 1)
                time.sleep(0.05) # Shorter delay

            server.quit()
            self.log_message("Disconnected from SMTP server.")

        except smtplib.SMTPAuthenticationError:
            err = "SMTP Authentication Error. Check Gmail & App Password. Use App Password for 2FA."
            self.log_message(err, error=True); messagebox.showerror("SMTP Auth Error", err)
            if not is_test: failed_count = len(emails_to_send_list) - sent_count
        except smtplib.SMTPConnectError:
            err = "SMTP Connection Error. Could not connect. Check internet."
            self.log_message(err, error=True); messagebox.showerror("SMTP Connection Error", err)
            if not is_test: failed_count = len(emails_to_send_list) - sent_count
        except Exception as e:
            self.log_message(f"An unexpected error during sending: {e}", error=True); messagebox.showerror("Sending Error", f"Unexpected error: {e}")
        finally:
            self.log_message(f"Process finished. Sent: {sent_count}, Failed: {failed_count if not is_test else 'N/A for test'}.")
            if hasattr(self, 'send_button'): self.send_button.config(state=tk.NORMAL)
            if hasattr(self, 'progress_bar') and not is_test and emails_to_send_list : self.progress_bar['value'] = self.progress_bar['maximum'] 

    def update_progress(self, current_step):
        if hasattr(self, 'progress_bar'):
            self.progress_bar['value'] = current_step
            if self.root and self.root.winfo_exists(): self.root.update_idletasks()


    def send_emails_process(self):
        if not self.csv_data: messagebox.showerror("Error", "No CSV data loaded."); return
        
        email_col_name = self.email_column_var.get()
        if not email_col_name or email_col_name == "Not Mapped" or email_col_name not in self.csv_headers:
            messagebox.showerror("Error", "Email column not selected/invalid. Check 'Column Mapping' tab."); return

        sender_email = self.smtp_email_var.get()
        sender_password = self.smtp_password_var.get()
        if not sender_email or not sender_password: messagebox.showerror("Error", "Gmail address or App Password not set in active profile."); return
        if not self._is_valid_email(sender_email): messagebox.showerror("Error", "Invalid sender Gmail address format."); return

        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return
        
        cv_path = self.cv_file_path.get() # Get CV path from active profile
        if cv_path and not os.path.exists(cv_path):
            if not messagebox.askyesno("CV Path Invalid", f"The CV path for the active profile is invalid:\n'{cv_path}'\nContinue without attaching any CV?"): return
            self.log_message("Proceeding without CV attachment (path invalid).", error=False)
            cv_path = None # Nullify to prevent attachment attempt
        elif cv_path and not cv_path.lower().endswith(".pdf"):
            messagebox.showerror("Error", "CV file must be a PDF. Please correct in active profile."); return
        elif not cv_path:
             self.log_message("No CV selected in active profile. Emails will be sent without attachments.", error=False)


        emails_to_send_list = []
        for i, row_data in enumerate(self.csv_data):
            recipient_email = row_data.get(email_col_name)
            if not recipient_email or not self._is_valid_email(recipient_email):
                self.log_message(f"Skipping row {i+1}: Invalid/missing email '{recipient_email}'.", error=True)
                continue
            
            current_subject = subject_template
            current_body = body_template
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                csv_col_for_placeholder = self.column_mappings[key].get()
                value_to_insert = ""
                if csv_col_for_placeholder and csv_col_for_placeholder != "Not Mapped" and csv_col_for_placeholder in row_data:
                    value_to_insert = row_data[csv_col_for_placeholder]
                current_subject = current_subject.replace(placeholder, str(value_to_insert))
                current_body = current_body.replace(placeholder, str(value_to_insert))
            
            emails_to_send_list.append({
                'recipient_email': recipient_email, 
                'subject': current_subject, 
                'body': current_body,
                'row_identifier': f"CSV Row {i+1}"
            })
        
        if not emails_to_send_list:
            messagebox.showinfo("Info", "No valid recipient emails found in CSV data to send."); return

        send_time_str = self.preferred_send_time_var.get()
        if not self._validate_send_time_format(send_time_str):
            messagebox.showerror("Error", "Invalid 'Preferred Send Time' format. Use HH:MM or leave blank."); return

        if send_time_str:
            try:
                now = datetime.datetime.now()
                hour, minute = map(int, send_time_str.split(':'))
                scheduled_dt = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                
                if scheduled_dt <= now: 
                     messagebox.showerror("Schedule Error", f"Scheduled time {send_time_str} is in the past for today. Choose a future time or clear to send now.")
                     return

                delay_ms = int((scheduled_dt - now).total_seconds() * 1000)
                if not messagebox.askyesno("Confirm Schedule", f"Schedule sending of {len(emails_to_send_list)} emails for {send_time_str} today? App must remain open."):
                    return
                
                self.log_message(f"Emails scheduled for {send_time_str}. Delaying for {delay_ms/1000:.0f}s.")
                self.root.after(delay_ms, lambda: self._perform_email_sending(emails_to_send_list))
            except ValueError:
                messagebox.showerror("Error", "Invalid time format in 'Preferred Send Time'. Use HH:MM."); return
        else: 
            if not messagebox.askyesno("Confirm Send", f"Send {len(emails_to_send_list)} emails now?"): return
            self._perform_email_sending(emails_to_send_list)


    def send_manual_email_process(self):
        recipient_email = self.manual_email_var.get()
        first_name = self.manual_first_name_var.get()
        company_name = self.manual_company_name_var.get()
        role = self.manual_role_var.get() 

        if not recipient_email or not self._is_valid_email(recipient_email):
            messagebox.showerror("Validation Error", "Valid recipient email is required for manual send."); return
        
        sender_email = self.smtp_email_var.get()
        sender_password = self.smtp_password_var.get()
        if not sender_email or not sender_password: messagebox.showerror("Error", "Gmail address or App Password not set in active profile."); return
        if not self._is_valid_email(sender_email): messagebox.showerror("Error", "Invalid sender Gmail address format."); return
        
        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return

        cv_path = self.cv_file_path.get() # CV from active profile
        if cv_path and not os.path.exists(cv_path):
            if not messagebox.askyesno("CV Path Invalid", f"The CV path for the active profile is invalid:\n'{cv_path}'\nContinue without attaching CV?"): return
            cv_path = None 
        elif cv_path and not cv_path.lower().endswith(".pdf"):
             messagebox.showerror("Error", "CV file must be a PDF."); return
        elif not cv_path:
             self.log_message("No CV selected in active profile for manual send.", error=False)


        current_subject = subject_template.replace(DEFAULT_PLACEHOLDERS["FIRST_NAME"], first_name or "")
        current_subject = current_subject.replace(DEFAULT_PLACEHOLDERS["COMPANY_NAME"], company_name or "")
        current_subject = current_subject.replace(DEFAULT_PLACEHOLDERS["ROLE"], role or "")
        current_subject = current_subject.replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], "") 

        current_body = body_template.replace(DEFAULT_PLACEHOLDERS["FIRST_NAME"], first_name or "")
        current_body = current_body.replace(DEFAULT_PLACEHOLDERS["COMPANY_NAME"], company_name or "")
        current_body = current_body.replace(DEFAULT_PLACEHOLDERS["ROLE"], role or "")
        current_body = current_body.replace(DEFAULT_PLACEHOLDERS["LAST_NAME"], "")

        email_details = [{'recipient_email': recipient_email, 'subject': current_subject, 'body': current_body, 'row_identifier': "Manual Send"}]
        
        if not messagebox.askyesno("Confirm Manual Send", f"Send email to {recipient_email} now?"): return
        self._perform_email_sending(email_details, is_test=True) 


    def send_test_email_process(self):
        sender_email = self.smtp_email_var.get()
        if not sender_email or not self._is_valid_email(sender_email):
            messagebox.showerror("Error", "Your (sender) Gmail address is not set or invalid in active profile."); return
        
        subject_template = self.email_subject_var.get()
        body_template = self.email_body_text_widget.get("1.0", tk.END) if self.email_body_text_widget else ""
        if not subject_template or not body_template.strip(): messagebox.showerror("Error", "Email subject or body empty in active profile."); return

        test_fill_data = {}
        active_tab_text = ""
        try:
            active_tab_text = self.notebook.tab(self.notebook.select(), "text")
        except tk.TclError: # If no tab is selected or notebook not fully ready
            pass


        if active_tab_text == "Manual Send" and (self.manual_first_name_var.get() or self.manual_company_name_var.get()): 
            self.log_message("Preparing test email using data from 'Manual Send' tab inputs.")
            test_fill_data[DEFAULT_PLACEHOLDERS["FIRST_NAME"]] = self.manual_first_name_var.get() or "[TEST_FIRST_NAME]"
            test_fill_data[DEFAULT_PLACEHOLDERS["LAST_NAME"]] = "" 
            test_fill_data[DEFAULT_PLACEHOLDERS["COMPANY_NAME"]] = self.manual_company_name_var.get() or "[TEST_COMPANY]"
            test_fill_data[DEFAULT_PLACEHOLDERS["ROLE"]] = self.manual_role_var.get() or "[TEST_ROLE]"
        elif self.csv_data:
            self.log_message("Preparing test email using data from the first CSV row.")
            first_row = self.csv_data[0]
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                csv_col_name = self.column_mappings[key].get()
                if csv_col_name and csv_col_name != "Not Mapped" and csv_col_name in first_row:
                    test_fill_data[placeholder] = first_row[csv_col_name]
                else:
                    test_fill_data[placeholder] = f"[{key.upper()}_TEST_DATA]"
        else:
            self.log_message("Preparing test email using generic placeholder data (no CSV/Manual data).")
            for key, placeholder in DEFAULT_PLACEHOLDERS.items():
                test_fill_data[placeholder] = f"[{key.upper()}_GENERIC_TEST]"

        current_subject = subject_template
        current_body = body_template
        for placeholder, value in test_fill_data.items():
            current_subject = current_subject.replace(placeholder, str(value))
            current_body = current_body.replace(placeholder, str(value))

        email_details = [{'recipient_email': sender_email, 'subject': f"[TEST EMAIL] {current_subject}", 'body': current_body, 'row_identifier': "Test Email"}]
        
        if not messagebox.askyesno("Confirm Test Send", f"Send a test email to yourself ({sender_email})?"): return
        self._perform_email_sending(email_details, is_test=True)


if __name__ == "__main__":
    try:
        root = tk.Tk()
        try:
            style = ttk.Style(root)
            available_themes = style.theme_names()
            # print("Available themes:", available_themes) 
            if 'clam' in available_themes: style.theme_use('clam')
            elif 'vista' in available_themes and os.name == 'nt': style.theme_use('vista')
            elif 'aqua' in available_themes and os.name == 'posix': style.theme_use('aqua') # macOS
            # Add other theme preferences if needed
        except Exception as e: print(f"Could not set theme: {e}")

        app = BulkEmailerApp(root)
        root.mainloop()
    except tk.TclError as e:
        # This error typically means no display is available
        print(f"Tkinter TclError: {e}")
        print("This Python script requires a graphical display environment to run.")
        print("Please ensure you are running this on a system with a display,")
        print("or if running via SSH, ensure X11 forwarding is enabled.")
    except Exception as e:
        print(f"An unexpected error occurred when starting the application: {e}")
        import traceback
        traceback.print_exc()

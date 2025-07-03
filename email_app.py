import streamlit as st
import pandas as pd
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import os
import re
import datetime
import time
import uuid
from io import StringIO

# --- Configuration and Constants ---
CONFIG_FILE = "bulk_emailer_config_profiles.json"
SCHEDULED_CAMPAIGNS_FILE = "scheduled_campaigns.json"

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

# --- Helper Functions for State Management and Logic ---

def get_default_profile_settings():
    """Returns a dictionary with default settings for a new profile."""
    return {
        "cv_file_path": "",
        "cv_file_name": "", # Store the name of the uploaded CV
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
        "smtp_email": "",
        "smtp_password": "",
        "enable_cc": False,
        "cc_email": ""
    }

def initialize_session_state():
    """Initializes session state variables if they don't exist."""
    if 'initialized' not in st.session_state:
        st.session_state.initialized = True
        st.session_state.profiles = {}
        st.session_state.active_profile_name = DEFAULT_PROFILE_NAME
        st.session_state.csv_headers = []
        st.session_state.csv_data = []
        st.session_state.custom_email_batch = []
        st.session_state.log_messages = []
        st.session_state.uploaded_csv = None
        st.session_state.uploaded_cv = None
        
        # Load existing config or create a default one
        load_app_config()

def log_message(message, error=False):
    """Adds a message to the log in session state."""
    log_entry = f"{'ERROR' if error else 'INFO'}: {message}"
    st.session_state.log_messages.insert(0, log_entry)
    # Keep log size manageable
    st.session_state.log_messages = st.session_state.log_messages[:50]

def save_app_config():
    """Saves all profiles and the active profile name to the config file."""
    app_config = {
        "active_profile_name": st.session_state.active_profile_name,
        "profiles": st.session_state.profiles
    }
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(app_config, f, indent=4)
        log_message("Application configuration (all profiles) saved.")
        st.success("Configuration saved successfully!")
    except Exception as e:
        log_message(f"Error saving application configuration: {e}", error=True)
        st.error(f"Failed to save configuration: {e}")

def load_app_config():
    """Loads profiles and active profile from the config file."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                app_config = json.load(f)
            st.session_state.profiles = app_config.get("profiles", {})
            st.session_state.active_profile_name = app_config.get("active_profile_name", DEFAULT_PROFILE_NAME)
        
        # Ensure default profile exists if no profiles were loaded
        if not st.session_state.profiles:
            st.session_state.profiles[DEFAULT_PROFILE_NAME] = get_default_profile_settings()
            st.session_state.active_profile_name = DEFAULT_PROFILE_NAME
        
        log_message("Application configuration loaded.")
    except Exception as e:
        log_message(f"Error loading config or config is corrupted: {e}. Creating a default.", error=True)
        st.session_state.profiles = {DEFAULT_PROFILE_NAME: get_default_profile_settings()}
        st.session_state.active_profile_name = DEFAULT_PROFILE_NAME

def get_active_profile():
    """Returns the data for the currently active profile."""
    return st.session_state.profiles.get(st.session_state.active_profile_name, get_default_profile_settings())

def save_active_profile_data(data):
    """Saves data to the active profile in session state."""
    st.session_state.profiles[st.session_state.active_profile_name] = data

def load_csv_data(uploaded_file):
    """Loads and parses a CSV file, auto-detecting column mappings."""
    if uploaded_file is None:
        st.session_state.csv_data = []
        st.session_state.csv_headers = []
        log_message("CSV file cleared.", error=True)
        return

    try:
        # To read the uploaded file, we need to treat it as a string
        string_data = StringIO(uploaded_file.getvalue().decode('utf-8'))
        # Use pandas for robust CSV parsing
        df = pd.read_csv(string_data)
        
        st.session_state.csv_headers = df.columns.tolist()
        st.session_state.csv_data = df.to_dict('records')

        log_message(f"Successfully loaded {len(st.session_state.csv_data)} rows from '{uploaded_file.name}'.")
        st.success(f"Loaded {len(st.session_state.csv_data)} rows from `{uploaded_file.name}`.")

        # Auto-detect columns
        auto_detect_columns()

    except Exception as e:
        log_message(f"Error reading CSV file: {e}", error=True)
        st.error(f"Error reading CSV file: {e}")
        st.session_state.csv_data = []
        st.session_state.csv_headers = []

def auto_detect_columns():
    """Automatically detects and sets column mappings based on common names."""
    profile = get_active_profile()
    mappings = profile["column_mappings"]
    detected_count = 0

    # Auto-detect email column
    if not profile.get("email_column"):
        for header in st.session_state.csv_headers:
            if any(pattern in header.lower() for pattern in AUTO_DETECT_PATTERNS["email_column"]):
                profile["email_column"] = header
                detected_count += 1
                break

    # Auto-detect placeholder columns
    for key, patterns in AUTO_DETECT_PATTERNS.items():
        if key == "email_column": continue
        if not mappings.get(key):
            for header in st.session_state.csv_headers:
                if any(pattern in header.lower() for pattern in patterns):
                    mappings[key] = header
                    detected_count += 1
                    break
    
    if detected_count > 0:
        log_message(f"Auto-detected {detected_count} column mappings.")
        st.info(f"Auto-detected {detected_count} column mappings. Please verify them in the 'Column Mapping' tab.")
    
    save_active_profile_data(profile)


def send_emails_process(email_list, subject, body, cv_attachment_bytes, cv_filename, smtp_details, cc_details):
    """The core process for sending emails."""
    if not smtp_details["email"] or not smtp_details["password"]:
        st.error("SMTP email and password are required in the 'Settings & Send' tab.")
        return

    total_emails = len(email_list)
    if total_emails == 0:
        st.warning("There are no emails to send. Load a CSV or add to the custom batch.")
        return

    progress_bar = st.progress(0)
    status_text = st.empty()
    success_count = 0
    fail_count = 0

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(smtp_details["email"], smtp_details["password"])
        log_message("SMTP server connected successfully.")

        for i, recipient in enumerate(email_list):
            try:
                # Validate recipient email
                if not re.match(r"[^@]+@[^@]+\.[^@]+", recipient.get("email", "")):
                    log_message(f"Skipping invalid email address: {recipient.get('email', 'N/A')}", error=True)
                    fail_count += 1
                    continue

                msg = MIMEMultipart()
                msg['From'] = smtp_details["email"]
                msg['To'] = recipient["email"]
                msg['Subject'] = subject.format(**recipient)

                # Add CC if enabled and valid
                if cc_details["enabled"] and cc_details["email"]:
                     msg['Cc'] = cc_details["email"]
                     to_addrs = [recipient["email"]] + [cc_details["email"]]
                else:
                     to_addrs = [recipient["email"]]

                # Attach body
                # Use .format() to replace placeholders like {{FIRST_NAME}}
                formatted_body = body.format(**recipient)
                msg.attach(MIMEText(formatted_body, 'plain'))

                # Attach CV if provided
                if cv_attachment_bytes and cv_filename:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(cv_attachment_bytes)
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename= {cv_filename}")
                    msg.attach(part)

                server.sendmail(smtp_details["email"], to_addrs, msg.as_string())
                log_message(f"Email sent to {recipient['email']}")
                success_count += 1

            except Exception as e:
                log_message(f"Failed to send to {recipient.get('email', 'N/A')}: {e}", error=True)
                fail_count += 1
            
            # Update progress
            progress = (i + 1) / total_emails
            progress_bar.progress(progress)
            status_text.text(f"Sent: {success_count} | Failed: {fail_count} | Total: {total_emails}")
            time.sleep(0.1) # Small delay to prevent overwhelming the server

        server.quit()
        log_message("SMTP server disconnected.")
        st.success(f"Email sending process completed. Sent: {success_count}, Failed: {fail_count}")

    except smtplib.SMTPAuthenticationError:
        st.error("SMTP Authentication Error: Check your email/password. You may need to use an 'App Password' for Gmail.")
        log_message("SMTP Authentication Error.", error=True)
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        log_message(f"An unexpected error occurred during sending: {e}", error=True)


# --- Streamlit UI ---

st.set_page_config(layout="wide", page_title="Streamlit Bulk Emailer")

st.title("✉️ Streamlit Bulk Emailer")
st.caption("A web-based tool to send personalized bulk emails using CSV files.")

# Initialize state on first run
initialize_session_state()

# --- Sidebar for Profile Management ---
with st.sidebar:
    st.header("👤 Profile Management")

    profile_names = list(st.session_state.profiles.keys())
    try:
        # Ensure active profile exists, otherwise default to first
        active_profile_index = profile_names.index(st.session_state.active_profile_name)
    except ValueError:
        active_profile_index = 0
        st.session_state.active_profile_name = profile_names[0] if profile_names else DEFAULT_PROFILE_NAME


    st.session_state.active_profile_name = st.selectbox(
        "Active Profile",
        profile_names,
        index=active_profile_index,
        help="Select the configuration profile to use."
    )

    st.markdown("---")
    
    with st.expander("Create or Delete Profiles"):
        new_profile_name = st.text_input("New Profile Name")
        if st.button("Create New Profile"):
            if new_profile_name:
                if new_profile_name in st.session_state.profiles:
                    st.warning(f"Profile '{new_profile_name}' already exists.")
                else:
                    st.session_state.profiles[new_profile_name] = get_default_profile_settings()
                    st.session_state.active_profile_name = new_profile_name
                    log_message(f"Profile '{new_profile_name}' created.")
                    st.success(f"Profile '{new_profile_name}' created and activated.")
                    st.rerun()
            else:
                st.warning("Please enter a name for the new profile.")

        if st.session_state.active_profile_name != DEFAULT_PROFILE_NAME:
            if st.button(f"Delete '{st.session_state.active_profile_name}' Profile", type="primary"):
                del st.session_state.profiles[st.session_state.active_profile_name]
                log_message(f"Profile '{st.session_state.active_profile_name}' deleted.")
                st.session_state.active_profile_name = DEFAULT_PROFILE_NAME
                st.success("Profile deleted.")
                st.rerun()

    if st.button("Save All Profile Settings"):
        save_app_config()

# Get a reference to the active profile's data
active_profile_data = get_active_profile()


# --- Main Interface with Tabs ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📂 Data & CV", 
    "↔️ Column Mapping", 
    "✍️ Email Content", 
    "⚙️ Settings & Send",
    "📊 Log"
])


with tab1:
    st.header("Load Data Source and CV")
    st.markdown("Upload a CSV file with recipient data and optionally attach a CV for the email campaign.")
    
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Recipient Data (CSV)")
        uploaded_csv = st.file_uploader(
            "Upload a CSV file",
            type=['csv'],
            help="The CSV should contain columns for email addresses and any other data you want to use as placeholders (e.g., FIRST_NAME, COMPANY_NAME)."
        )
        if uploaded_csv:
            if st.session_state.get('uploaded_csv_name') != uploaded_csv.name:
                 st.session_state.uploaded_csv_name = uploaded_csv.name
                 load_csv_data(uploaded_csv)
        
        if st.session_state.csv_data:
            st.info(f"**{len(st.session_state.csv_data)}** rows loaded. Preview:")
            st.dataframe(pd.DataFrame(st.session_state.csv_data).head())

    with col2:
        st.subheader("CV Attachment (Optional)")
        uploaded_cv = st.file_uploader(
            "Upload your CV",
            type=['pdf', 'doc', 'docx'],
            help="This file will be attached to every email sent in the campaign."
        )
        if uploaded_cv:
            # Store CV bytes and name in session state for later use
            st.session_state.uploaded_cv_bytes = uploaded_cv.getvalue()
            active_profile_data['cv_file_name'] = uploaded_cv.name
            save_active_profile_data(active_profile_data)
            st.success(f"CV `{uploaded_cv.name}` is ready to be attached.")
        elif active_profile_data.get('cv_file_name'):
             st.info(f"Using previously uploaded CV: `{active_profile_data['cv_file_name']}`")


with tab2:
    st.header("Map CSV Columns to Placeholders")
    st.markdown("Match the columns from your CSV file to the email placeholders. The system will try to auto-detect them.")

    if not st.session_state.csv_headers:
        st.warning("Please upload a CSV file in the 'Data & CV' tab to see mapping options.")
    else:
        options = [""] + st.session_state.csv_headers
        
        # Email Column
        try:
            email_index = options.index(active_profile_data.get("email_column", ""))
        except ValueError:
            email_index = 0
        active_profile_data["email_column"] = st.selectbox(
            "**Email Column** (Required)",
            options,
            index=email_index,
            help="Select the column containing recipient email addresses."
        )

        st.markdown("---")
        st.subheader("Placeholder Mapping")
        
        mappings = active_profile_data.get("column_mappings", {})
        
        for key, placeholder in DEFAULT_PLACEHOLDERS.items():
            try:
                current_mapping_index = options.index(mappings.get(key, ""))
            except ValueError:
                current_mapping_index = 0
            
            mappings[key] = st.selectbox(
                f"`{placeholder}`",
                options,
                index=current_mapping_index,
                key=f"map_{key}"
            )
        
        active_profile_data["column_mappings"] = mappings
        save_active_profile_data(active_profile_data)

with tab3:
    st.header("Compose Your Email")
    st.markdown("Write the subject and body of your email. Use the placeholders from the mapping tab.")
    
    st.subheader("Email Subject")
    active_profile_data["email_subject"] = st.text_input(
        "Subject Line",
        value=active_profile_data.get("email_subject", ""),
        help="Use placeholders like {{COMPANY_NAME}}."
    )

    st.subheader("Email Body")
    active_profile_data["email_body"] = st.text_area(
        "Body Content",
        value=active_profile_data.get("email_body", ""),
        height=300,
        help="Use placeholders like {{FIRST_NAME}}. Note: This is a plain text editor."
    )
    save_active_profile_data(active_profile_data)


with tab4:
    st.header("Configure & Send")
    st.markdown("Enter your SMTP server details and start the email campaign.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("SMTP Settings (Gmail)")
        st.info("For Gmail, you may need to generate an 'App Password'. [Learn more](https://support.google.com/accounts/answer/185833)", icon="💡")
        
        active_profile_data["smtp_email"] = st.text_input(
            "Your Gmail Address",
            value=active_profile_data.get("smtp_email", "")
        )
        active_profile_data["smtp_password"] = st.text_input(
            "Your Gmail App Password",
            type="password",
            value=active_profile_data.get("smtp_password", "")
        )

    with col2:
        st.subheader("CC Settings (Optional)")
        active_profile_data["enable_cc"] = st.checkbox(
            "Enable CC",
            value=active_profile_data.get("enable_cc", False)
        )
        if active_profile_data["enable_cc"]:
            active_profile_data["cc_email"] = st.text_input(
                "CC Email Address",
                value=active_profile_data.get("cc_email", "")
            )

    save_active_profile_data(active_profile_data)

    st.markdown("---")
    st.subheader("Launch Campaign")
    
    if st.button("🚀 Send Emails", type="primary", use_container_width=True):
        with st.spinner("Preparing to send emails..."):
            # Prepare data for the sending function
            email_list = st.session_state.csv_data
            
            # Create a dictionary of substitutions for each recipient
            email_data_list = []
            mappings = active_profile_data.get("column_mappings", {})
            email_col = active_profile_data.get("email_column")

            if not email_col:
                st.error("Email column is not set in the 'Column Mapping' tab.")
            else:
                for row in email_list:
                    recipient_data = {"email": row.get(email_col)}
                    for key, col_name in mappings.items():
                        if col_name:
                            recipient_data[key] = row.get(col_name, '')
                    email_data_list.append(recipient_data)

                # Get other details
                subject = active_profile_data.get("email_subject", "")
                body = active_profile_data.get("email_body", "")
                cv_bytes = st.session_state.get("uploaded_cv_bytes")
                cv_name = active_profile_data.get("cv_file_name")
                smtp_details = {
                    "email": active_profile_data.get("smtp_email"),
                    "password": active_profile_data.get("smtp_password")
                }
                cc_details = {
                    "enabled": active_profile_data.get("enable_cc"),
                    "email": active_profile_data.get("cc_email")
                }

                send_emails_process(email_data_list, subject, body, cv_bytes, cv_name, smtp_details, cc_details)

with tab5:
    st.header("Activity Log")
    st.markdown("Shows the most recent actions and errors from the application.")
    
    if st.button("Clear Log"):
        st.session_state.log_messages = []
        st.rerun()

    log_container = st.container(height=400)
    for msg in st.session_state.log_messages:
        if "ERROR" in msg:
            log_container.error(msg)
        else:
            log_container.info(msg)


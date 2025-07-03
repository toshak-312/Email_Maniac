import streamlit as st
import pandas as pd
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re

# =================================================================================
# 1. App Title
# =================================================================================
st.set_page_config(page_title="Toshak's Bulk Deployer", layout="wide")
st.title("Toshak's Bulk Deployer for Outreach 🚀")

# =================================================================================
# Helper Functions
# =================================================================================

def send_email(smtp_server, smtp_port, smtp_user, smtp_pass, from_addr, to_addr, subject, body, signature, attachment_bytes=None, attachment_name=None):
    """
    Connects to the SMTP server and sends a single email.
    """
    try:
        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = subject

        # =========================================================================
        # 4. HTML Email Support & 5. Signature Support
        # =========================================================================
        # Combine body and signature, both supporting HTML
        full_body = body + "<br><br>" + signature
        msg.attach(MIMEText(full_body, 'html'))

        # =========================================================================
        # 3. Custom Attachment Support
        # =========================================================================
        if attachment_bytes and attachment_name:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_bytes)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {attachment_name}")
            msg.attach(part)

        # Send the email
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def render_template(template, data_row, neutral_tokens, column_mappings):
    """
    Replaces neutral tokens in a template string with data from a CSV row.
    """
    rendered = template
    for token in neutral_tokens:
        mapped_column = column_mappings.get(token)
        if mapped_column and mapped_column in data_row and pd.notna(data_row[mapped_column]):
            rendered = rendered.replace(f"{{{{{token}}}}}", str(data_row[mapped_column]))
        else:
            # If a token is not mapped or the data is missing, remove it
            rendered = rendered.replace(f"{{{{{token}}}}}", "")
    return rendered

# =================================================================================
# 2. Neutral Terminology for Variable Fields
# =================================================================================
NEUTRAL_TOKENS = ["A_01", "A_02", "A_03", "A_04", "A_05"]

# =================================================================================
# Sidebar for Configuration
# =================================================================================
with st.sidebar:
    st.header("⚙️ Configuration")

    st.subheader("SMTP Credentials")
    smtp_server = st.text_input("SMTP Server", "smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)
    smtp_user = st.text_input("Your Email Address")
    smtp_pass = st.text_input("Your App Password", type="password")

    st.subheader("Defaults")
    # =========================================================================
    # 3. Custom Attachment Support
    # =========================================================================
    st.markdown("---")
    uploaded_attachment = st.file_uploader("Attach a file to all emails")
    attachment_bytes = None
    attachment_name = None
    if uploaded_attachment:
        attachment_bytes = uploaded_attachment.getvalue()
        attachment_name = uploaded_attachment.name
        st.success(f"Attachment '{attachment_name}' ready!")

    # =========================================================================
    # 5. Signature Support
    # =========================================================================
    st.markdown("---")
    st.subheader("Signature")
    signature = st.text_area("Paste your HTML signature here", height=200)
    if signature:
        st.write("Signature Preview:")
        st.markdown(signature, unsafe_allow_html=True)


# =================================================================================
# Main App Body - Using Tabs for different modes
# =================================================================================

tab1, tab2 = st.tabs(["📤 Bulk Send (from CSV)", "✉️ Manual Send"])

# =================================================================================
# TAB 1: Bulk Send from CSV
# =================================================================================
with tab1:
    st.header("1. Compose Your Email")
    
    col1, col2 = st.columns(2)
    with col1:
        email_subject_template = st.text_input("Email Subject Template")

    with col2:
        # A little trick to show placeholder text
        st.write("Available Tokens:")
        st.code(f"{{{{{', '.join(NEUTRAL_TOKENS)}}}}}", language="text")

    email_body_template = st.text_area("Email Body Template", height=300)
    st.info("💡 HTML supported: use <b>bold</b>, <br> for line breaks, <a href='...'>links</a>, etc.")


    st.header("2. Upload & Map Your Data")
    uploaded_csv = st.file_uploader("Upload your recipient list (CSV)", type="csv")

    column_mappings = {}
    if uploaded_csv:
        try:
            df = pd.read_csv(uploaded_csv)
            st.write("CSV Preview:")
            st.dataframe(df.head())
            
            st.subheader("Map CSV Columns to Tokens")
            st.warning("Map the recipient's email address and any tokens you used in your template.")

            cols = st.columns(len(NEUTRAL_TOKENS) + 1)
            
            # Mapping for Email
            with cols[0]:
                email_column = st.selectbox("Email Column", options=df.columns, index=None, placeholder="Select Email Column")
            
            # Mapping for Neutral Tokens
            for i, token in enumerate(NEUTRAL_TOKENS):
                with cols[i+1]:
                    column_mappings[token] = st.selectbox(f"{{{{{token}}}}}", options=df.columns, index=None, placeholder=f"Map {token}")

        except Exception as e:
            st.error(f"Failed to read CSV: {e}")
            df = None
    
    st.header("3. Review & Send")
    if st.button("🚀 Send Bulk Emails", disabled=(uploaded_csv is None or 'df' not in locals())):
        if not all([smtp_server, smtp_port, smtp_user, smtp_pass]):
            st.error("SMTP credentials are required in the sidebar.")
        elif not email_column:
            st.error("Please map the 'Email Column' before sending.")
        else:
            total_emails = len(df)
            st.info(f"Starting bulk deployment to {total_emails} recipients...")
            
            progress_bar = st.progress(0)
            log_messages = []
            
            for i, row in df.iterrows():
                recipient_email = row[email_column]
                
                # Simple regex for email validation
                if not re.match(r"[^@]+@[^@]+\.[^@]+", recipient_email):
                    log_messages.append(f"Skipped: Invalid email format - {recipient_email}")
                    continue
                    
                # Render subject and body
                subject = render_template(email_subject_template, row, NEUTRAL_TOKENS, column_mappings)
                body = render_template(email_body_template, row, NEUTRAL_TOKENS, column_mappings)

                # Send email
                success, error_msg = send_email(
                    smtp_server, smtp_port, smtp_user, smtp_pass, 
                    smtp_user, recipient_email, subject, body, signature, 
                    attachment_bytes, attachment_name
                )
                
                if success:
                    log_messages.append(f"✅ Sent to: {recipient_email}")
                else:
                    log_messages.append(f"❌ Failed for: {recipient_email} | Error: {error_msg}")
                
                # Update progress bar
                progress_bar.progress((i + 1) / total_emails)
                time.sleep(0.1) # Small delay to prevent overwhelming the server

            st.success("Bulk deployment finished!")
            
            # Display logs
            with st.expander("View Send Log"):
                for msg in log_messages:
                    if "✅" in msg:
                        st.success(msg)
                    else:
                        st.error(msg)


# =================================================================================
# TAB 2: Manual Send
# =================================================================================
with tab2:
    st.header("Send a Single, Ad-hoc Email")
    st.info("This section allows you to send a single email without needing a CSV file. It uses the same templates, signature, and attachment from the other tabs and sidebar.")

    recipient_email_manual = st.text_input("Recipient Email Address")
    
    st.subheader("Provide values for the neutral tokens:")
    manual_token_values = {}
    cols_manual = st.columns(len(NEUTRAL_TOKENS))
    for i, token in enumerate(NEUTRAL_TOKENS):
        with cols_manual[i]:
            manual_token_values[token] = st.text_input(f"Value for {{{{{token}}}}}", key=f"manual_{token}")

    if st.button("🚀 Send Manual Email"):
        if not all([smtp_server, smtp_port, smtp_user, smtp_pass]):
            st.error("SMTP credentials are required in the sidebar.")
        elif not re.match(r"[^@]+@[^@]+\.[^@]+", recipient_email_manual):
            st.error("Please enter a valid recipient email address.")
        else:
            # Create a mock data row from manual inputs
            mock_row = {f"{{{token}}}": val for token, val in manual_token_values.items()}
            
            # This is a simplified render for manual mode
            def render_manual(template, values):
                rendered = template
                for token, value in values.items():
                    rendered = rendered.replace(f"{{{{{token}}}}}", value)
                return rendered

            subject = render_manual(st.session_state.get('email_subject_template', ''), manual_token_values)
            body = render_manual(st.session_state.get('email_body_template', ''), manual_token_values)
            
            st.info(f"Sending to {recipient_email_manual}...")
            
            success, error_msg = send_email(
                smtp_server, smtp_port, smtp_user, smtp_pass,
                smtp_user, recipient_email_manual, subject, body, signature,
                attachment_bytes, attachment_name
            )

            if success:
                st.success(f"✅ Successfully sent email to {recipient_email_manual}!")
            else:
                st.error(f"❌ Failed to send email. Error: {error_msg}")

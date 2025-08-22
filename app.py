import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
import os
import tempfile
import zipfile
from io import BytesIO
import re
import platform
import subprocess
import pythoncom

if platform.system() == 'Windows':
    import win32com.client
import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
import requests
from urllib.parse import urlparse
import time
import threading
import gspread
from gspread_dataframe import get_as_dataframe
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(
    page_title="SLM Generator & Email Bulk Tool",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .feature-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #007bff;
        margin-bottom: 1rem;
    }
    .status-success { color: #28a745; }
    .status-error { color: #dc3545; }
    .status-warning { color: #ffc107; }
</style>
""", unsafe_allow_html=True)

if 'email_templates' not in st.session_state:
    st.session_state.email_templates = {}
if 'email_list' not in st.session_state:
    st.session_state.email_list = []
if 'sending_emails' not in st.session_state:
    st.session_state.sending_emails = False

def main():
    st.markdown("""
    <div class="main-header">
        <h1>🚀 SLM Generator & Email Bulk Tool</h1>
        <p>All-in-one solution untuk generate SLM dan bulk email automation</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("🎯 Menu Utama")
        feature = st.radio(
            "Pilih Fitur:",
            ["📄 SLM Generator", "📧 Email Bulk Automation"],
            index=0
        )
        
        st.markdown("---")
        st.markdown("### ℹ️ Info")
        if feature == "📄 SLM Generator":
            st.info("📋 Generate dokumen SLM dari data Excel dan template Word")
        else:
            st.info("✉️ Kirim email massal dengan template dan attachment otomatis")
    
    if feature == "📄 SLM Generator":
        render_slm_generator()
    else:
        render_email_bulk()

# ========== SLM GENERATOR FUNCTIONS ==========

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF with multiple fallback methods"""
    try:
        # Method 1: Try docx2pdf (works on some systems)
        try:
            from docx2pdf import convert
            # Initialize COM for Windows
            if platform.system() == "Windows":
                pythoncom.CoInitialize()
            convert(docx_path, pdf_path)
            if platform.system() == "Windows":
                pythoncom.CoUninitialize()
            return True
        except Exception as e:
            st.warning(f"docx2pdf failed: {str(e)}, trying alternative method...")
            
        # Method 2: Try using win32com directly (Windows only)
        if platform.system() == "Windows":
            try:
                pythoncom.CoInitialize()
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
                doc.Close()
                word.Quit()
                pythoncom.CoUninitialize()
                return True
            except Exception as e:
                st.warning(f"Win32com failed: {str(e)}, trying LibreOffice...")
                if platform.system() == "Windows":
                    pythoncom.CoUninitialize()
        
        # Method 3: Try LibreOffice (cross-platform)
        try:
            # Check if LibreOffice is available
            libreoffice_paths = [
                "libreoffice",
                "/usr/bin/libreoffice",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
            ]
            
            libreoffice_cmd = None
            for path in libreoffice_paths:
                try:
                    subprocess.run([path, "--version"], capture_output=True, timeout=10)
                    libreoffice_cmd = path
                    break
                except:
                    continue
            
            if libreoffice_cmd:
                output_dir = os.path.dirname(pdf_path)
                cmd = [
                    libreoffice_cmd,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", output_dir,
                    docx_path
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                
                # LibreOffice creates PDF with same name as DOCX but with .pdf extension
                expected_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
                
                if os.path.exists(expected_pdf):
                    if expected_pdf != pdf_path:
                        os.rename(expected_pdf, pdf_path)
                    return True
                else:
                    st.warning(f"LibreOffice conversion failed: {result.stderr}")
            else:
                st.warning("LibreOffice not found on system")
                
        except Exception as e:
            st.warning(f"LibreOffice method failed: {str(e)}")
        
        # Method 4: Return DOCX file if PDF conversion fails
        st.error("⚠️ PDF conversion failed. Will provide DOCX files instead.")
        return False
        
    except Exception as e:
        st.error(f"All conversion methods failed: {str(e)}")
        return False

def validate_filename_prefix(prefix):
    """Validasi format filename prefix (contoh: VI-001)"""
    pattern = r'^[IVX]+[-]\d{3}$'
    return bool(re.match(pattern, prefix))

def process_slm_data(df, template_file, filename_prefix):
    """Proses data SLM dan generate dokumen"""
    
    month_id = {
        'January': 'Januari', 'February': 'Februari', 'March': 'Maret',
        'April': 'April', 'May': 'Mei', 'June': 'Juni',
        'July': 'Juli', 'August': 'Agustus', 'September': 'September',
        'October': 'Oktober', 'November': 'November', 'December': 'Desember'
    }
    df_copy = df.copy()
    
    if 'SLMDate' in df_copy.columns:
        df_copy['SLMDate'] = pd.to_datetime(df_copy['SLMDate']).dt.strftime('%d %B %Y')
        for en_month, id_month in month_id.items():
            df_copy['SLMDate'] = df_copy['SLMDate'].str.replace(en_month, id_month)    
    if 'PaidDate' in df_copy.columns:
        df_copy['PaidDate'] = pd.to_datetime(df_copy['PaidDate']).dt.strftime('%d %B %Y')
        for en_month, id_month in month_id.items():
            df_copy['PaidDate'] = df_copy['PaidDate'].str.replace(en_month, id_month)    
    if 'ClientPhone' in df_copy.columns:
        df_copy['ClientPhone'] = df_copy['ClientPhone'].astype(str).str.zfill(12)    
    if 'Outstanding' in df_copy.columns:
        df_copy['Outstanding'] = df_copy['Outstanding'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))    
    if 'PaidAmount' in df_copy.columns:
        df_copy['PaidAmount'] = df_copy['PaidAmount'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

    roman_month, seq_str = filename_prefix.split('-')
    sequence_number = int(seq_str)

    temp_dir = tempfile.mkdtemp()
    output_files = []
    conversion_success = True

    try:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, (_, row) in enumerate(df_copy.iterrows()):
            progress = (idx + 1) / len(df_copy)
            progress_bar.progress(progress)
            status_text.text(f"Memproses SLM {idx + 1} dari {len(df_copy)}: {row.get('ClientName', 'Unknown')}")
            
            doc = DocxTemplate(template_file)
            context = row.to_dict()
            doc.render(context)

            seq_num_str = str(sequence_number).zfill(3)
            client_name = str(row.get('ClientName', 'Unknown')).upper().replace('/', '-')
            platform = str(row.get('PlatformName', 'Unknown')).upper()
            
            filename = f"SLM {roman_month} {seq_num_str} {client_name} ({platform})"
            temp_docx_path = os.path.join(temp_dir, f"{filename}.docx")
            temp_pdf_path = os.path.join(temp_dir, f"{filename}.pdf")

            doc.save(temp_docx_path)
            pdf_success = convert_docx_to_pdf(temp_docx_path, temp_pdf_path)

            if pdf_success and os.path.exists(temp_pdf_path):
                output_files.append(temp_pdf_path)
                try:
                    os.remove(temp_docx_path)
                except:
                    pass
            else:
                output_files.append(temp_docx_path)
                conversion_success = False
            sequence_number += 1

        progress_bar.progress(1.0)
        if conversion_success:
            status_text.text("✅ Semua SLM berhasil diproses sebagai PDF!")
        else:
            status_text.text("⚠️ Beberapa SLM diproses sebagai DOCX karena konversi PDF gagal")
        
        return output_files, temp_dir, conversion_success
        
    except Exception as e:
        st.error(f"Error saat memproses SLM: {str(e)}")
        return None, None, False

def create_zip_file(pdf_files):
    """Buat file ZIP berisi semua PDF"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for pdf_path in pdf_files:
            filename = os.path.basename(pdf_path)
            zip_file.write(pdf_path, filename)
    
    zip_buffer.seek(0)
    return zip_buffer

def render_slm_generator():
    st.markdown("## 📄 SLM Generator")
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 Upload Files")
        
        excel_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="File Excel berisi data SLM",
            key="slm_excel"
        )
        
        template_file = st.file_uploader(
            "Upload Template Word",
            type=['docx'],
            help="Template dokumen Word untuk SLM",
            key="slm_template"
        )
        
        st.markdown("### ⚙️ Pengaturan")
        filename_prefix = st.text_input(
            "Filename Prefix",
            placeholder="VI-001",
            help="Format: ROMAWI-NOMOR (contoh: VI-001)",
            key="slm_prefix"
        )
        
        if filename_prefix and not validate_filename_prefix(filename_prefix):
            st.error("❌ Format filename prefix salah! Gunakan format seperti: VI-001")
        
        if excel_file is not None:
            try:
                df = pd.read_excel(excel_file)
                st.markdown("### 📊 Preview Data Excel")
                st.dataframe(df.head(10), use_container_width=True)
                st.info(f"📈 Total records: {len(df)} | Columns: {', '.join(df.columns.tolist())}")
                
            except Exception as e:
                st.error(f"❌ Error membaca file Excel: {str(e)}")
                df = None
        else:
            st.info("👆 Upload file Excel untuk melihat preview data")
            df = None
    
    with col2:
        st.markdown("### ℹ️ Status")
        
        status_items = [
            ("Excel File", excel_file is not None),
            ("Template Word", template_file is not None),
            ("Filename Prefix", filename_prefix and validate_filename_prefix(filename_prefix))
        ]
        
        for item, status in status_items:
            if status:
                st.success(f"✅ {item}")
            else:
                st.error(f"❌ {item}")
    
    st.markdown("---")
    
    if st.button("🚀 Generate SLM", type="primary", use_container_width=True, key="generate_slm"):
        if not excel_file:
            st.error("❌ Harap upload file Excel terlebih dahulu!")
            return        
        if not template_file:
            st.error("❌ Harap upload template Word terlebih dahulu!")
            return
        if not filename_prefix or not validate_filename_prefix(filename_prefix):
            st.error("❌ Harap masukkan filename prefix dengan format yang benar!")
            return
        
        with st.spinner("⏳ Sedang memproses SLM..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_template:
                temp_template.write(template_file.read())
                temp_template_path = temp_template.name
            
            try:
                output_files, temp_dir, conversion_success = process_slm_data(df, temp_template_path, filename_prefix)
                
                if output_files:
                    zip_buffer = create_zip_file(output_files)
                    file_extension = "PDF" if conversion_success else "DOCX"
                    
                    if conversion_success:
                        st.success(f"🎉 Berhasil menggenerate {len(output_files)} file PDF!")
                    else:
                        st.warning(f"⚠️ Berhasil menggenerate {len(output_files)} file DOCX (PDF conversion gagal)")
                        st.info("💡 **Solusi untuk PDF conversion:**\n"
                               "1. Install LibreOffice (recommended): https://www.libreoffice.org/\n"
                               "2. Install Microsoft Office (Windows)\n"
                               "3. Gunakan DOCX files dan convert manual")
                    
                    st.download_button(
                        label=f"⬇️ Download All SLM ({file_extension})",
                        data=zip_buffer,
                        file_name=f"SLM_{filename_prefix.replace('-', '_')}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                    
                    os.unlink(temp_template_path)
                    if temp_dir and os.path.exists(temp_dir):
                        import shutil
                        shutil.rmtree(temp_dir)
                
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {str(e)}")
                if os.path.exists(temp_template_path):
                    os.unlink(temp_template_path)


# ========== EMAIL BULK FUNCTIONS ==========

def get_gspread_client():
    """Setup Google Sheets client"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"❌ GSpread Auth Error: {str(e)}")
        return None

def load_spreadsheet_data(sheet_url):
    """Load data from Google Spreadsheet"""
    try:
        client = get_gspread_client()
        if not client:
            return False, "Failed to authenticate with Google Sheets"
        
        sheet = client.open_by_url(sheet_url)
        
        try:
            email_worksheet = sheet.worksheet("EmailList")
            email_df = get_as_dataframe(email_worksheet).dropna(subset=["EmailAddress"])
            load_email_list_from_df(email_df)
            email_count = len(st.session_state.email_list)
        except Exception as e:
            return False, f"Failed to read 'EmailList' sheet: {str(e)}"

        try:
            template_worksheet = sheet.worksheet("BodySubject")
            template_df = get_as_dataframe(template_worksheet).dropna(subset=["Template Name"])
            load_templates_from_df(template_df)
            template_count = len(st.session_state.email_templates)
        except Exception as e:
            return False, f"Failed to read 'BodySubject' sheet: {str(e)}"

        return True, f"Successfully loaded {template_count} templates and {email_count} emails"
        
    except Exception as e:
        return False, f"Failed to load spreadsheet: {str(e)}"

def load_templates_from_df(df):
    """Load email templates from dataframe"""
    st.session_state.email_templates.clear()
    for _, row in df.iterrows():
        template_name = str(row.get('Template Name', ''))
        subject = str(row.get('Subject', ''))
        body = str(row.get('Body', ''))

        if template_name and template_name != 'nan':
            st.session_state.email_templates[template_name] = {
                'subject': subject,
                'body': body
            }

def load_email_list_from_df(df):
    """Load email list from dataframe"""
    st.session_state.email_list.clear()
    for _, row in df.iterrows():
        email = str(row.get('EmailAddress', ''))
        client_name = str(row.get('Client Name', ''))
        attachment = str(row.get('Attachment (for SLM)', ''))

        if email and email != 'nan' and '@' in email:
            st.session_state.email_list.append({
                'email': email,
                'client_name': client_name if client_name != 'nan' else '',
                'attachment': attachment if attachment != 'nan' else ''
            })

def download_attachment(url):
    """Download attachment from URL"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        parsed_url = urlparse(url)
        filename = os.path.basename(parsed_url.path) or "attachment"
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, filename)

        with open(temp_path, 'wb') as f:
            f.write(response.content)

        return temp_path
    except Exception as e:
        st.error(f"Failed to download attachment from {url}: {str(e)}")
        return None

def create_email_message(recipient_email, client_name, subject, body, attachment_url=None, include_attachments=False, cc_emails=None, bcc_emails=None):
    """Create email message"""
    final_subject = subject.replace('[Client Name]', client_name)
    final_body = body.replace('[Client Name]', client_name)

    msg = MIMEMultipart()
    msg['From'] = st.session_state.sender_email
    msg['To'] = recipient_email
    msg['Subject'] = final_subject
    msg['Date'] = formatdate(localtime=True)
    
    if cc_emails:
        msg['Cc'] = cc_emails
    if bcc_emails:
        msg['Bcc'] = bcc_emails
    
    msg.attach(MIMEText(final_body, 'plain'))

    if attachment_url and include_attachments:
        attachment_path = download_attachment(attachment_url)
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(attachment_path)}'
            )
            msg.attach(part)

            try:
                os.remove(attachment_path)
            except OSError:
                pass
    
    return msg

def send_email_smtp(msg, recipient_email, server, port, email, password):
    """Send email via SMTP"""
    try:
        all_recipients = [recipient_email]
        
        if msg['Cc']:
            cc_list = [email.strip() for email in msg['Cc'].split(',')]
            all_recipients.extend(cc_list)
        
        if msg['Bcc']:
            bcc_list = [email.strip() for email in msg['Bcc'].split(',')]
            all_recipients.extend(bcc_list)
        
        try:
            with smtplib.SMTP(server, int(port)) as smtp_server:
                smtp_server.starttls()
                smtp_server.login(email, password)
                smtp_server.sendmail(email, all_recipients, msg.as_string())
            return True
        except:
            with smtplib.SMTP_SSL(server, 465) as smtp_server:
                smtp_server.login(email, password)
                smtp_server.sendmail(email, all_recipients, msg.as_string())
            return True
            
    except Exception as e:
        st.error(f"SMTP failed for {recipient_email}: {str(e)}")
        return False

def render_email_bulk():
    st.markdown("## 📧 Email Bulk Automation")
    st.markdown("---")
    
    st.markdown("### 📊 Data Loading")
    col1, col2 = st.columns([3, 1])
    
    with col1:
        sheet_url = st.text_input(
            "Google Spreadsheet URL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="Paste URL Google Spreadsheet yang berisi data EmailList dan BodySubject"
        )
    
    with col2:
        if st.button("📥 Load Data", type="secondary"):
            if sheet_url:
                with st.spinner("Loading spreadsheet data..."):
                    success, message = load_spreadsheet_data(sheet_url)
                    if success:
                        st.success(message)
                    else:
                        st.error(message)
            else:
                st.error("Please enter a valid spreadsheet URL")
    
    if st.session_state.email_templates and st.session_state.email_list:
        st.success(f"✅ Data loaded: {len(st.session_state.email_templates)} templates, {len(st.session_state.email_list)} emails")
    elif st.session_state.email_templates or st.session_state.email_list:
        st.warning(f"⚠️ Data partially loaded: {len(st.session_state.email_templates)} templates, {len(st.session_state.email_list)} emails")
    else:
        st.info("ℹ️ No data loaded. Please load spreadsheet data first.")
    
    st.markdown("---")
    
    st.markdown("### ⚙️ Email Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        protocol = st.selectbox("Protocol", ["SMTP"], key="email_protocol")
        server = st.text_input("Server", value="smtp.gmail.com", key="email_server")
        port = st.number_input("Port", value=587, key="email_port")
        
    with col2:
        sender_email = st.text_input("Sender Email", key="email_sender")
        sender_password = st.text_input("Password", type="password", key="email_password")
        
        st.session_state.sender_email = sender_email
    
    st.markdown("#### 📨 Additional Recipients (Optional)")
    col1, col2 = st.columns(2)
    
    with col1:
        cc_emails = st.text_input(
            "CC Emails", 
            placeholder="email1@domain.com, email2@domain.com",
            help="Separate multiple emails with commas",
            key="cc_emails"
        )
    
    with col2:
        bcc_emails = st.text_input(
            "BCC Emails", 
            placeholder="email1@domain.com, email2@domain.com",
            help="Separate multiple emails with commas",
            key="bcc_emails"
        )
    
    st.markdown("---")
    
    st.markdown("### 📝 Template Selection")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.session_state.email_templates:
            selected_template = st.selectbox(
                "Select Template",
                list(st.session_state.email_templates.keys()),
                key="selected_email_template"
            )
        else:
            st.warning("⚠️ No templates available. Please load data first.")
            selected_template = None
    
    with col2:
        include_attachments = st.checkbox("Include Attachments", key="include_attachments")
    
    if selected_template:
        st.markdown("#### 👁️ Template Preview")
        template = st.session_state.email_templates[selected_template]
        
        col1, col2 = st.columns([1, 3])
        with col1:
            st.text("Subject:")
        with col2:
            st.code(template['subject'].replace('[Client Name]', '[Client Name]'), language=None)
        
        col1, col2 = st.columns([1, 3])
        with col1:
            st.text("Body:")
        with col2:
            st.text_area("", template['body'].replace('[Client Name]', '[Client Name]'), height=150, disabled=True, key="preview_body")
    
    st.markdown("---")
    
    st.markdown("### 🚀 Send Emails")
    
    can_send = all([
        st.session_state.email_templates,
        st.session_state.email_list,
        selected_template,
        sender_email,
        sender_password,
        server,
        port
    ])
    
    if not can_send:
        st.error("❌ Please fill all required fields and load data first.")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        if st.button(
            f"📧 Send {len(st.session_state.email_list)} Emails" if st.session_state.email_list else "📧 Send Emails",
            type="primary",
            disabled=not can_send,
            use_container_width=True,
            key="send_emails_btn"
        ):
            with st.expander("🔍 Email Send Summary", expanded=True):
                st.write(f"**Template:** {selected_template}")
                st.write(f"**Recipients:** {len(st.session_state.email_list)} emails")
                st.write(f"**Protocol:** {protocol}")
                st.write(f"**Attachments:** {'Yes' if include_attachments else 'No'}")
                if cc_emails:
                    st.write(f"**CC:** {cc_emails}")
                if bcc_emails:
                    st.write(f"**BCC:** {bcc_emails}")
                
                confirm_send = st.button("✅ Confirm & Send", type="primary", key="confirm_send")
                
                if confirm_send:
                    template = st.session_state.email_templates[selected_template]
                    total_emails = len(st.session_state.email_list)
                    sent_count = 0
                    failed_count = 0
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    log_container = st.container()
                    
                    with log_container:
                        st.markdown("#### 📝 Email Sending Log")
                        log_placeholder = st.empty()
                        log_messages = []
                    
                    for i, email_data in enumerate(st.session_state.email_list):
                        recipient = email_data['email']
                        client_name = email_data['client_name']
                        
                        status_text.text(f"Sending {i+1}/{total_emails}: {recipient}...")
                        
                        msg = create_email_message(
                            recipient,
                            client_name,
                            template['subject'],
                            template['body'],
                            email_data['attachment'] if include_attachments else None,
                            include_attachments,
                            cc_emails if cc_emails else None,
                            bcc_emails if bcc_emails else None
                        )
                        
                        success = send_email_smtp(msg, recipient, server, port, sender_email, sender_password)
                        
                        if success:
                            sent_count += 1
                            log_messages.append(f"✅ Successfully sent to {recipient}")
                        else:
                            failed_count += 1
                            log_messages.append(f"❌ Failed to send to {recipient}")
                        
                        progress_bar.progress((i + 1) / total_emails)
                        
                        log_placeholder.text_area(
                            "Sending Progress:",
                            "\n".join(log_messages[-10:]),  # Show last 10 messages
                            height=200,
                            disabled=True
                        )
                        
                        time.sleep(1)

                    status_text.text("✅ Email blast completed!")
                    
                    if failed_count > 0:
                        st.warning(f"⚠️ Completed with {failed_count} failures out of {total_emails} emails.")
                    else:
                        st.success(f"🎉 All {sent_count} emails sent successfully!")

if __name__ == "__main__":
    main()
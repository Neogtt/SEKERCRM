import streamlit as st
import pandas as pd
import io, os, re
from email.message import EmailMessage
import smtplib
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =========================
# Config
# =========================
st.set_page_config(page_title="Fair Mailer (BCC)", layout="wide")
EXCEL_FILE_ID = "1IF6CN4oHEMk6IEE40ZGixPkfnNHLYXnQ"  # Drive'daki Excel ID
LOCAL_FALLBACK = "D:/APP/temp.xlsx"                   # Lokal fallback
SHEET_NAME = "FuarMusteri"

FROM_EMAIL = "todo@sekeroglugroup.com"
FROM_PASS  = "vbgvforwwbcpzhxf"  # Google App Password (SMTP)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT_SSL = 465

# =========================
# Login (Boss only)
# =========================
USERS = {"Boss": "Seker12345!"}
if "user" not in st.session_state:
    st.session_state.user = None

def login_screen():
    st.title("Fair Mailer (BCC)")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Sign in"):
        if u in USERS and p == USERS[u]:
            st.session_state.user = u
            st.rerun()
        else:
            st.error("Invalid credentials.")

if not st.session_state.user:
    login_screen()
    st.stop()

st.title("📧 Fair Mailer — BCC bulk sender")

# =========================
# Drive client
# =========================
@st.cache_resource
def build_drive():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

drive_svc = build_drive()

def download_excel_from_drive(file_id: str, to_path: str = "fuar_temp.xlsx") -> str | None:
    try:
        req = drive_svc.files().get_media(fileId=file_id)
        with io.FileIO(to_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                status, done = downloader.next_chunk()
        return to_path
    except Exception as e:
        st.warning(f"Could not download from Drive: {e}")
        return None

# =========================
# Load data
# =========================
excel_path = download_excel_from_drive(EXCEL_FILE_ID, "fuar_temp.xlsx")
if excel_path is None or not os.path.exists(excel_path):
    excel_path = LOCAL_FALLBACK
    st.info(f"Using local fallback: {excel_path}")

try:
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
except Exception as e:
    st.error(f"Could not read sheet '{SHEET_NAME}': {e}")
    st.stop()

# Beklenen kolonlar: A: Fuar Adı, E: E-mail
# Esneklik için kolon adlarını indeksle de alalım:
def get_col(df, idx, default_name):
    try:
        return df.columns[idx]
    except:
        return default_name

col_fuar = "Fuar Adı" if "Fuar Adı" in df.columns else get_col(df, 0, "Fuar Adı")
col_email = "E-mail" if "E-mail" in df.columns else get_col(df, 4, "E-mail")

# Fuar listesi
fuar_list = sorted([str(x).strip() for x in df[col_fuar].dropna().unique() if str(x).strip() != ""])
fuar = st.selectbox("Select Fair", fuar_list)

# =========================
# Email utilities
# =========================
EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def clean_emails(series) -> list[str]:
    emails = []
    for v in series.dropna():
        s = str(v).strip()
        if not s:
            continue
        # bir hücrede çoklu e-posta varsa ayıralım: ; , boşluk
        parts = re.split(r"[;, \n]+", s)
        for p in parts:
            p = p.strip()
            if p and EMAIL_REGEX.match(p):
                emails.append(p)
    # unique, orijinal sırayı mümkün olduğunca koru
    seen = set()
    uniq = []
    for e in emails:
        if e.lower() not in seen:
            uniq.append(e)
            seen.add(e.lower())
    return uniq

# =========================
# Filter by fair & build recipient list
# =========================
rec_df = df[df[col_fuar].astype(str).str.strip() == fuar] if fuar else df.iloc[0:0]
recipients = clean_emails(rec_df[col_email]) if not rec_df.empty else []

st.markdown(f"*Selected fair:* {fuar}")
st.write(f"Recipients found: *{len(recipients)}*")
if recipients:
    st.dataframe(pd.DataFrame({"Recipients (BCC)": recipients}), use_container_width=True, height=210)
else:
    st.info("No recipients for this fair.")

# =========================
# Compose
# =========================
default_subject = f"Thank you for visiting us at {fuar}"
default_body = (
    f"Dear Partner,\n\n"
    f"It was a pleasure meeting you at {fuar}. Please find attached our materials.\n\n"
    f"Best regards,\n"
    f"Sekeroglu Export Team"
)

st.subheader("Compose Email")
subject = st.text_input("Subject (English only)", value=default_subject)
body = st.text_area("Body (English only)", value=default_body, height=220)

st.subheader("Attachments (optional)")
files = st.file_uploader("Select one or more files to attach", type=None, accept_multiple_files=True)

# =========================
# Send
# =========================
colA, colB = st.columns([1,1])
with colA:
    confirm = st.checkbox(f"I confirm to send to {len(recipients)} recipients via BCC.", value=False)
with colB:
    send_btn = st.button("Send Email", type="primary", use_container_width=True)

def send_bcc_email(from_email: str, password: str, subject: str, body: str, bcc_list: list[str], attachments=None):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = from_email               # BCC gönderimde To kendimiz olsun
    # BCC list normal header'a eklenmez; SMTP çağrısında ayrı verilir.
    msg.set_content(body)

    # Attachments
    if attachments:
        for f in attachments:
            data = f.read()
            fname = f.name
            maintype = "application"
            subtype = "octet-stream"
            # İçerik türünü Streamlit'ten alamadığımız için generic ekliyoruz
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=fname)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT_SSL) as smtp:
        smtp.login(from_email, password)
        smtp.sendmail(from_email, [from_email] + bcc_list, msg.as_string())

if send_btn:
    if not recipients:
        st.error("No recipients to send.")
    elif not subject.strip() or not body.strip():
        st.error("Subject and Body are required.")
    elif not confirm:
        st.warning("Please confirm before sending.")
    else:
        try:
            send_bcc_email(FROM_EMAIL, FROM_PASS, subject, body, recipients, attachments=files)
            st.success(f"Email sent successfully to {len(recipients)} recipients (BCC).")
        except Exception as e:
            st.error(f"Send failed: {e}")

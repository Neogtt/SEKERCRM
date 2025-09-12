# email_app.py
import streamlit as st
import pandas as pd
import io, os
import smtplib
from pathlib import Path
from email.message import EmailMessage
from email.utils import make_msgid
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ===========================
# ==== APP CONFIG
# ===========================
st.set_page_config(page_title="ŞEKEROĞLU • Fair Mailer", layout="wide")

# ---- Drive'daki Excel ve lokal yedek
EXCEL_FILE_ID = "1IF6CN4oHEMk6IEE40ZGixPkfnNHLYXnQ"  # Google Drive'daki Excel (aynı dosyayı CRM'inizde kullanıyorsunuz)
LOCAL_FILE    = "D:/APP/temp.xlsx"                    # Lokal yedek (Drive erişilemezse)

# ---- Gönderici bilgileri
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 465
FROM_EMAIL    = "export1@sekeroglugroup.com"
FROM_PASSWORD = "vbgvforwwbcpzhxf"   # Gmail App Password (talebiniz üzerine koda gömüldü)

# ===========================
# ==== GOOGLE DRIVE HELPER
# ===========================
@st.cache_resource
def build_drive():
    """Service Account ile Drive istemcisi oluşturur."""
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def download_excel_from_drive(file_id: str, out_path: str) -> bool:
    """Drive'dan Excel indirir; başarıyla indirilirse True döner."""
    try:
        svc = build_drive()
        req = svc.files().get_media(fileId=file_id)
        with io.FileIO(out_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                status, done = downloader.next_chunk()
        return True
    except Exception as e:
        st.warning(f"Drive download failed, using local file if present. ({e})")
        return False

# ===========================
# ==== DATA LOAD
# ===========================
def load_fuar_sheet() -> pd.DataFrame:
    tmp_path = "fuar_temp.xlsx"
    ok = download_excel_from_drive(EXCEL_FILE_ID, tmp_path)
    use_path = tmp_path if ok and os.path.exists(tmp_path) else LOCAL_FILE

    if not os.path.exists(use_path):
        st.error("Neither Drive file nor local fallback could be loaded.")
        return pd.DataFrame()

    try:
        df = pd.read_excel(use_path, sheet_name="FuarMusteri")
    except Exception as e:
        st.error(f"Couldn't read 'FuarMusteri' sheet: {e}")
        return pd.DataFrame()

    # Beklenen sütunlar: A=Fuar Adı, E=E-mail (diğerleri opsiyonel)
    # Sütun adları değişik olabilir; güvenli şekilde adresi E sütunu olarak alalım:
    if df.shape[1] >= 5:
        if "E-mail" not in df.columns:
            df.rename(columns={df.columns[4]: "E-mail"}, inplace=True)
        if "Fuar Adı" not in df.columns:
            df.rename(columns={df.columns[0]: "Fuar Adı"}, inplace=True)

    # Email ve fuar adı boş/NaN olanları ayıkla
    df["E-mail"] = df["E-mail"].astype(str).str.strip()
    df["Fuar Adı"] = df["Fuar Adı"].astype(str).str.strip()
    df = df[(df["E-mail"] != "") & (df["Fuar Adı"] != "")]
    return df

df_fuar = load_fuar_sheet()

# ===========================
# ==== SIGNATURE (H. POLAT)
# ===========================
def build_signature_html() -> tuple[str, str | None]:
    """
    Hüseyin Polat imzası (HTML). Eğer logo.png varsa inline CID döner.
    return: (html, logo_cid_without_brackets_or_None)
    """
    logo_path = Path("logo.png")
    logo_cid = make_msgid() if logo_path.exists() else None
    logo_cid_trim = logo_cid[1:-1] if logo_cid else None

    logo_img_html = f'<img src="cid:{logo_cid_trim}" alt="Şekeroğlu A.Ş." style="width:180px;"><br><br>' if logo_cid_trim else ""

    html = f"""
    <br><br>
    <div style="font-family: Arial, sans-serif; font-size: 12px; color: #333;">
        {logo_img_html}
        <b style="color:#97B900; font-size:14px;">Huseyin POLAT</b><br>
        Export Sales Representative<br><br>
        📞 +90 531 765 69 60<br>
        ☎️ +90 850 420 27 00<br>
        ✉️ <a href="mailto:export1@sekeroglugroup.com">export1@sekeroglugroup.com</a><br>
        🌐 <a href="https://www.sekeroglugroup.com">www.sekeroglugroup.com</a><br><br>
        <b style="color:#97B900;">Şekeroğlu A.Ş.</b><br>
        Sanayi Mah. 60129 Nolu Cad. No:7 27110 Şehitkamil / Gaziantep<br><br>
        <small style="color:gray;">
          This email and its contents are confidential and intended only for the recipient.
          Unauthorized sharing or disclosure is strictly prohibited. If you received this message in error,
          please delete it immediately.
        </small>
    </div>
    """
    return html, logo_cid_trim

# ===========================
# ==== SMTP SEND
# ===========================
def send_one_email(to_addr: str, bcc_list: list[str], subject: str, body_html: str, attachments: list[bytes], attachment_names: list[str], logo_cid: str | None):
    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["To"] = to_addr
    if bcc_list:
        msg["Bcc"] = ", ".join(bcc_list)
    msg["Subject"] = subject
    msg.set_content(body_html, subtype="html")

    # Inline logo
    logo_path = Path("logo.png")
    if logo_cid and logo_path.exists():
        with open(logo_path, "rb") as img:
            msg.get_payload()[0].add_related(img.read(), maintype="image", subtype="png", cid=f"<{logo_cid}>")

    # Attachments
    for data, name in zip(attachments, attachment_names):
        msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=name)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.login(FROM_EMAIL, FROM_PASSWORD)
        smtp.send_message(msg)

# ===========================
# ==== UI
# ===========================
st.title("✉️ Şekeroğlu • Fair Email Sender")

if df_fuar.empty:
    st.stop()

# Fuar seçimi
fuarlar = sorted(df_fuar["Fuar Adı"].dropna().unique().tolist())
colA, colB = st.columns([2, 3])
with colA:
    fuar_sec = st.selectbox("Select a Fair (Fuar Adı):", fuarlar, index=0 if fuarlar else None)

filtered = df_fuar[df_fuar["Fuar Adı"] == fuar_sec].copy() if fuarlar else pd.DataFrame()
emails = sorted(filtered["E-mail"].dropna().unique().tolist())

with colB:
    st.markdown(f"**Recipients for '{fuar_sec}':** {len(emails)}")

st.dataframe(pd.DataFrame({"Recipients": emails}), use_container_width=True, height=240)

st.markdown("---")

# Mail içeriği
subject = st.text_input("Subject (English only):", value="Follow-up from Şekeroğlu A.Ş. – Thank you for visiting our booth")
default_body = (
    "Dear Partner,<br><br>"
    "Thank you for visiting our booth during the fair. We would be pleased to support your inquiries regarding our products and pricing.<br><br>"
    "Please feel free to reply to this email for any questions.<br><br>"
    "Best regards,"
)
body = st.text_area("Message (HTML allowed, English only):", value=default_body, height=220)

# BCC
bcc_text = st.text_input("BCC (comma separated):", value="export1@sekeroglugroup.com")

# Ekler
uploaded_files = st.file_uploader("Attachments (any file type, multiple allowed)", type=None, accept_multiple_files=True)

# Gönder
send_btn = st.button("🚀 Send Emails")

if send_btn:
    if not emails:
        st.error("No recipients found for the selected fair.")
    elif not subject.strip() or not body.strip():
        st.error("Subject and message cannot be empty.")
    else:
        signature_html, logo_cid = build_signature_html()
        final_body = body + signature_html

        # Prepare attachments
        attachments_data = []
        attachment_names = []
        if uploaded_files:
            for uf in uploaded_files:
                attachments_data.append(uf.getvalue())
                attachment_names.append(uf.name)

        bcc_list = [e.strip() for e in bcc_text.split(",") if e.strip()]

        success, fail = 0, 0
        for addr in emails:
            try:
                send_one_email(
                    to_addr=addr,
                    bcc_list=bcc_list,
                    subject=subject.strip(),
                    body_html=final_body,
                    attachments=attachments_data,
                    attachment_names=attachment_names,
                    logo_cid=logo_cid
                )
                success += 1
            except Exception as e:
                st.write(f"❌ {addr}: {e}")
                fail += 1

        st.success(f"Done. Sent: {success} | Failed: {fail}")
        if success:
            st.info("Tip: To improve deliverability, avoid very large attachments and consider sending in batches if the list is big.")

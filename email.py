import streamlit as st
import pandas as pd
import io, os, re
from email.message import EmailMessage
import smtplib

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ===========================
# ==== AYARLAR
# ===========================
st.set_page_config(page_title="Fuar E-Posta Gönderici", layout="wide")

EXCEL_FILE_ID = "1IF6CN4oHEMk6IEE40ZGixPkfnNHLYXnQ"  # Drive'daki Excel ID
LOCAL_FALLBACK = "D:/APP/temp.xlsx"   # Lokal yedek (opsiyonel)

# ===========================
# ==== GOOGLE DRIVE
# ===========================
@st.cache_resource
def build_drive():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

drive_svc = build_drive()

def download_excel_file(file_id: str, local_path: str = "fuar_temp.xlsx") -> str | None:
    try:
        req = drive_svc.files().get_media(fileId=file_id)
        with io.FileIO(local_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                status, done = downloader.next_chunk()
        return local_path
    except Exception as e:
        st.warning(f"Drive'dan indirme başarısız: {e}")
        return None

# ===========================
# ==== VERİYİ YÜKLE
# ===========================
st.title("🎫 Fuar E-Posta Gönderici")

excel_path = download_excel_file(EXCEL_FILE_ID, "fuar_temp.xlsx")
if (excel_path is None or not os.path.exists("fuar_temp.xlsx")) and os.path.exists(LOCAL_FALLBACK):
    excel_path = LOCAL_FALLBACK
    st.info(f"Lokal dosya kullanılıyor: {excel_path}")

if not excel_path or not os.path.exists(excel_path):
    st.error("Excel dosyası bulunamadı. ID ve yetkileri kontrol edin.")
    st.stop()

try:
    # FuarMusteri sayfası: A sütunu = Fuar Adı, E sütunu = E-mail
    df = pd.read_excel(excel_path, sheet_name="FuarMusteri")
except Exception as e:
    st.error(f"FuarMusteri sayfası okunamadı: {e}")
    st.stop()

# E-mail sütunu (E sütunu) ve fuar adı (A sütunu) esnek yakalama
# Kullanıcı özelindeki kolon başlıkları farklı olabilir diye indeks bazlı da destekleyelim.
def pick_col_by_index_or_name(frame: pd.DataFrame, idx: int, fallback_names: list[str]) -> pd.Series:
    try:
        s = frame.iloc[:, idx]
        s.name = s.name or f"col_{idx}"
        return s
    except Exception:
        for name in fallback_names:
            if name in frame.columns:
                return frame[name]
        # Hiçbiri yoksa boş seri
        return pd.Series(dtype=object)

col_fuar = pick_col_by_index_or_name(df, 0, ["Fuar Adı", "FuarAdi", "Fuar"])
col_mail = pick_col_by_index_or_name(df, 4, ["E-mail", "Email", "E posta", "E_posta", "E-Mail"])

# Temel tabloyu normalize et
work = pd.DataFrame({
    "Fuar Adı": col_fuar.astype(str).str.strip(),
    "E-mail": col_mail.astype(str).str.strip(),
})
# Geçerli e-mail filtresi
email_pat = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
work = work[work["E-mail"].str.match(email_pat, na=False)]

fuar_list = sorted([x for x in work["Fuar Adı"].dropna().unique() if x])
secili_fuar = st.selectbox("Fuar adı seçin", fuar_list, index=0 if fuar_list else None)

if not fuar_list:
    st.info("FuarMusteri sayfasında 'Fuar Adı' verisi bulunamadı.")
    st.stop()

filtered = work[work["Fuar Adı"] == secili_fuar].copy()
alici_listesi = sorted(filtered["E-mail"].dropna().unique().tolist())

st.markdown(f"*Seçilen fuar:* {secili_fuar}")
st.markdown(f"*Alıcı sayısı:* {len(alici_listesi)}")

with st.expander("Alıcıları Göster"):
    st.write(pd.DataFrame({"E-mail": alici_listesi}))

# ===========================
# ==== E-POSTA GÖNDERİM ARAYÜZÜ
# ===========================
st.markdown("---")
st.subheader("E-posta İçeriği")

varsayilan_konu = f"{secili_fuar} Hakkında Bilgilendirme"
varsayilan_govde = (
    f"Merhaba,\n\n{secili_fuar} kapsamında görüştüğümüz için teşekkür ederiz. "
    "Aşağıda ürün ve hizmetlerimize dair kısa bilgileri bulabilirsiniz.\n\n"
    "Sorularınız için bu e-posta üzerinden bize dönebilirsiniz.\n\nSaygılarımızla,\nŞekeroğlu İhracat"
)

kol1, kol2 = st.columns(2)
with kol1:
    konu = st.text_input("Konu", value=varsayilan_konu)
with kol2:
    gonderici_isim = st.text_input("Gönderici Adı", value="Şekeroğlu İhracat")

govde = st.text_area("Mesaj", value=varsayilan_govde, height=220)

tek_tek_gonder = st.checkbox("Alıcılara tek tek gönder (önerilir)", value=True)
test_modu = st.checkbox("Önce test olarak sadece bana gönder", value=False)
test_mail = st.text_input("Test mail adresi", value="")

st.markdown("*Gönderim ayarları* st.secrets['smtp'] içinde tanımlanmalıdır:")
st.code(
    """# .streamlit/secrets.toml
[gcp_service_account]
# ... service account JSON içeriği ...

[smtp]
from_email = "todo@sekeroglugroup.com"
password   = "uygulama_şifresi_veya_smtp_parolası"
host       = "smtp.gmail.com"
port       = 465
""",
    language="toml"
)

def send_email(to_addr: str, subject: str, body: str, from_email: str, password: str, host: str, port: int, sender_name: str | None = None):
    msg = EmailMessage()
    frm = f"{sender_name} <{from_email}>" if sender_name else from_email
    msg["Subject"] = subject
    msg["From"] = frm
    msg["To"] = to_addr
    msg.set_content(body)

    with smtplib.SMTP_SSL(host, port) as smtp:
        smtp.login(from_email, password)
        smtp.send_message(msg)

def get_smtp_secrets():
    try:
        cfg = st.secrets["smtp"]
        from_email = cfg.get("from_email")
        password = cfg.get("password")
        host = cfg.get("host", "smtp.gmail.com")
        port = int(cfg.get("port", 465))
        if not from_email or not password:
            raise KeyError("from_email / password eksik.")
        return from_email, password, host, port
    except Exception as e:
        st.error(f"SMTP ayarları eksik veya hatalı: {e}")
        return None

st.markdown("---")
gonder_btn = st.button("📨 E-postaları Gönder", type="primary", disabled=(len(alici_listesi) == 0))

if gonder_btn:
    if not konu.strip() or not govde.strip():
        st.error("Konu ve mesaj boş olamaz.")
    else:
        smtp_cfg = get_smtp_secrets()
        if smtp_cfg is None:
            st.stop()
        from_email, password, host, port = smtp_cfg

        try:
            if test_modu:
                if not test_mail or not re.match(email_pat, test_mail):
                    st.error("Geçerli bir test mail adresi girin.")
                    st.stop()
                send_email(test_mail, konu, govde, from_email, password, host, port, gonderici_isim)
                st.success(f"✅ Test e-postası gönderildi: {test_mail}")
            else:
                if tek_tek_gonder:
                    basarili, hatali = 0, 0
                    for addr in alici_listesi:
                        try:
                            send_email(addr, konu, govde, from_email, password, host, port, gonderici_isim)
                            basarili += 1
                        except Exception:
                            hatali += 1
                    st.success(f"✅ Gönderim tamamlandı. Başarılı: {basarili}, Hatalı: {hatali}")
                else:
                    # Tek e-postada BCC ile
                    msg = EmailMessage()
                    frm = f"{gonderici_isim} <{from_email}>" if gonderici_isim else from_email
                    msg["Subject"] = konu
                    msg["From"] = frm
                    # 'To' alanına kendinizi koyun, alıcıları BCC yapalım
                    msg["To"] = from_email
                    msg["Bcc"] = ", ".join(alici_listesi)
                    msg.set_content(govde)
                    with smtplib.SMTP_SSL(host, port) as smtp:
                        smtp.login(from_email, password)
                        smtp.send_message(msg)
                    st.success(f"✅ Tek mail + BCC ile gönderildi. Alıcı sayısı: {len(alici_listesi)}")
        except Exception as e:
            st.error(f"Gönderim hatası: {e}")

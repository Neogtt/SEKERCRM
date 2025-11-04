import streamlit as st
import pandas as pd
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import io, os, datetime, tempfile, re, json, time, uuid, html
import numpy as np
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import streamlit.components.v1 as components
import matplotlib.pyplot as plt

st.set_page_config(page_title="ŞEKEROĞLU İHRACAT CRM", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700&display=swap');

    :root {
        color-scheme: dark;
    }

    body, .stApp {
        background-color: #0f1722;
        color: #e2e8f0;
        font-family: "Manrope", "Inter", "Segoe UI", sans-serif;
    }

    header[data-testid="stHeader"] {
        background: rgba(15, 23, 42, 0.4);
        backdrop-filter: blur(18px);
        border-bottom: 1px solid rgba(148, 163, 184, 0.25);
    }

    header[data-testid="stHeader"] * {
        color: #e2e8f0 !important;
    }

    section[data-testid="stSidebar"] {
        background: radial-gradient(circle at top, rgba(15, 23, 42, 0.95), #0b1120 70%);
        border-right: 1px solid rgba(148, 163, 184, 0.18);
    }

    section[data-testid="stSidebar"] > div {
        padding-top: 1.5rem;
        padding-bottom: 1.5rem;
    }

    div[data-testid="stSidebar"] .stButton > button {
        background: rgba(15, 23, 42, 0.8);
        border: 1px solid rgba(148, 163, 184, 0.35);
        color: #e2e8f0;
        border-radius: 10px;
        padding: 0.5rem 0.75rem;
        width: 100%;
        transition: all 0.2s ease;
    }

    div[data-testid="stSidebar"] .stButton > button:hover {
        border-color: rgba(56, 189, 248, 0.65);
        color: #38bdf8;
        background: rgba(30, 41, 59, 0.95);
    }

    div.block-container {
        padding-top: 2.5rem;
        padding-bottom: 2.5rem;
    }

    div.block-container, div[data-testid="stVerticalBlock"] {
        color: #e2e8f0;
    }

    .stApp main .block-container {
        max-width: 1200px;
    }

    .stMarkdown, .stText, .stDataFrame {
        color: inherit !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

EMBED_IMAGES = True

CURRENCY_SYMBOLS = ["USD", "$", "€", "EUR", "₺", "TL", "tl", "Tl"]

ETA_COLUMNS = ["Müşteri Adı", "Proforma No", "Sevk Tarihi", "ETA Tarihi", "Açıklama"]

def smart_to_num(value):
    if pd.isna(value):
        return 0.0

    sanitized = str(value).strip()
    for symbol in CURRENCY_SYMBOLS:
        sanitized = sanitized.replace(symbol, "")

    sanitized = sanitized.replace("\u00A0", "").replace(" ", "")

    try:
        return float(sanitized)
    except Exception:
        pass

    if "," in sanitized:
        try:
            return float(sanitized.replace(".", "").replace(",", "."))
        except Exception:
            pass

    return 0.0


def güvenli_sil(path, tekrar=5, bekle=1):
    for _ in range(tekrar):
        try:
            os.remove(path)
            return True
        except PermissionError:
            time.sleep(bekle)
        except FileNotFoundError:
            return True
    return False

# ==== KULLANICI GİRİŞİ SİSTEMİ ====
USERS = {
    "export1": "Seker12345!",
    "admin": "Seker12345!",
    "Boss": "Seker12345!",
    "Muhammed": "Seker12345!",
}
if "user" not in st.session_state:
    st.session_state.user = None
if "sync_status" not in st.session_state:
    st.session_state.sync_status = None
if "_selected_fuar" not in st.session_state:
    st.session_state["_selected_fuar"] = ""
    
def login_screen():
    st.title("ŞEKEROĞLU CRM - Giriş Ekranı")
    username = st.text_input("Kullanıcı Adı")
    password = st.text_input("Şifre", type="password")
    if st.button("Giriş Yap"):
        if username in USERS and password == USERS[username]:
            st.session_state.user = username
            st.success("Giriş başarılı!")
            st.rerun()
        else:
            st.error("Kullanıcı adı veya şifre hatalı.")

if not st.session_state.user:
    login_screen()
    st.stop()

# Sol menüde çıkış
if st.sidebar.button("Çıkış Yap"):
    st.session_state.user = None
    st.rerun()

def _request_manual_sync():
    st.session_state["_sync_requested"] = True


if st.session_state.sync_status:
    status_type, status_msg = st.session_state.sync_status
    display_fn = getattr(st, status_type, st.info)
    display_fn(status_msg)

st.sidebar.button("🔁 Excel Senkronizasyonu", on_click=_request_manual_sync)

# --- Referans listeler ---
ulke_listesi = sorted([
    "Afganistan","Almanya","Amerika Birleşik Devletleri","Andorra","Angola","Antigua ve Barbuda","Arjantin",
    "Arnavutluk","Avustralya","Avusturya","Azerbaycan","Bahamalar","Bahreyn","Bangladeş","Barbados","Belçika",
    "Belize","Benin","Beyaz Rusya","Bhutan","Birleşik Arap Emirlikleri","Birleşik Krallık","Bolivya",
    "Bosna-Hersek","Botsvana","Brezilya","Brunei","Bulgaristan","Burkina Faso","Burundi","Butan",
    "Cezayir","Çad","Çekya","Çin","Danimarka","Doğu Timor","Dominik Cumhuriyeti","Dominika","Ekvador",
    "Ekvator Ginesi","El Salvador","Endonezya","Eritre","Ermenistan","Estonya","Etiyopya","Fas",
    "Fiji","Fildişi Sahili","Filipinler","Filistin","Finlandiya","Fransa","Gabon","Gambia",
    "Gana","Gine","Gine-Bissau","Grenada","Guatemala","Guyana","Güney Afrika","Güney Kore",
    "Güney Sudan","Gürcistan","Haiti","Hindistan","Hırvatistan","Hollanda","Honduras","Hong Kong",
    "Irak","İran","İrlanda","İspanya","İsrail","İsveç","İsviçre","İtalya","İzlanda","Jamaika",
    "Japonya","Kamboçya","Kamerun","Kanada","Karadağ","Katar","Kazakistan","Kenya","Kırgızistan",
    "Kiribati","Kolombiya","Komorlar","Kongo","Kongo Demokratik Cumhuriyeti","Kostarika","Küba",
    "Kuveyt","Kuzey Kore","Kuzey Makedonya","Laos","Lesotho","Letonya","Liberya","Libya",
    "Liechtenstein","Litvanya","Lübnan","Lüksemburg","Macaristan","Madagaskar","Malavi","Maldivler",
    "Malezya","Mali","Malta","Marshall Adaları","Meksika","Mısır","Mikronezya","Moğolistan","Moldova",
    "Monako","Morityus","Mozambik","Myanmar","Namibya","Nauru","Nepal","Nijer","Nijerya",
    "Nikaragua","Norveç","Orta Afrika Cumhuriyeti","Özbekistan","Pakistan","Palau","Panama","Papua Yeni Gine",
    "Paraguay","Peru","Polonya","Portekiz","Romanya","Ruanda","Rusya","Saint Kitts ve Nevis",
    "Saint Lucia","Saint Vincent ve Grenadinler","Samoa","San Marino","Sao Tome ve Principe","Senegal",
    "Seyşeller","Sırbistan","Sierra Leone","Singapur","Slovakya","Slovenya","Solomon Adaları","Somali",
    "Sri Lanka","Sudan","Surinam","Suriye","Suudi Arabistan","Svaziland","Şili","Tacikistan","Tanzanya",
    "Tayland","Tayvan","Togo","Tonga","Trinidad ve Tobago","Tunus","Tuvalu","Türkiye","Türkmenistan",
    "Uganda","Ukrayna","Umman","Uruguay","Ürdün","Vanuatu","Vatikan","Venezuela","Vietnam",
    "Yemen","Yeni Zelanda","Yunanistan","Zambiya","Zimbabve"
]) + ["Diğer"]

temsilci_listesi = ["KEMAL İLKER ÇELİKKALKAN", "HÜSEYİN POLAT", "EFE YILDIRIM", "FERHAT ŞEKEROĞLU"]

# --- Sabitler ---
LOGO_FILE_ID     = "1DCxtSsAeR7Zfk2IQU0UMGmD0uTdNO1B3"
LOGO_LOCAL_NAME  = "logo1.png"
EXCEL_FILE_ID    = "1C8OpNAIRySkWYTI9jBaboV-Rq85UbVD9"
EVRAK_KLASOR_ID  = "14FTE1oSeIeJ6Y_7C0oQyZPKC8dK8hr1J"
FIYAT_TEKLIFI_ID = "1TNjwx-xhmlxNRI3ggCJA7jaCAu9Lt_65"

COUNTRY_LANGUAGE_MAP = {
    "Türkiye": "tr",
    "Amerika Birleşik Devletleri": "en",
    "Birleşik Krallık": "en",
    "Kanada": "en",
    "Avustralya": "en",
    "Almanya": "de",
    "Avusturya": "de",
    "İsviçre": "de",
    "Fransa": "fr",
    "Belçika": "fr",
    "İspanya": "es",
    "Meksika": "es",
    "Kolombiya": "es",
    "Arjantin": "es",
    "Birleşik Arap Emirlikleri": "ar",
    "Suudi Arabistan": "ar",
    "Katar": "ar",
    "Kuveyt": "ar",
}

FAIR_MAIL_TEMPLATES = {
    "tr": {
        "subject": "Fuar Görüşmemiz Hakkında",
        "body": (
            "Merhaba,\n\n"
            "Fuarda standımızı ziyaret ettiğiniz için teşekkür ederiz. Sunduğumuz ürün ve"
            " çözümler hakkında sorularınızı yanıtlamaktan memnuniyet duyarız."
            "\n\n"
            "İhtiyaçlarınızı daha iyi anlayabilmek ve iş birliği fırsatlarını görüşmek için"
            " uygun olduğunuz bir zamanı paylaşmanızı rica ederiz.\n\n"
            "Saygılarımızla,\nŞekeroğlu İhracat Ekibi"
        ),
    },
    "en": {
        "subject": "Thank You for Visiting Şekeroğlu at the Fair",
        "body": (
            "Hello,\n\n"
            "Thank you for taking the time to meet with us during the trade fair. We would"
            " be delighted to continue the conversation and share tailored solutions for"
            " your business.\n\n"
            "Please let us know a convenient time for a follow-up call or meeting so that"
            " we can discuss the next steps together.\n\n"
            "Best regards,\nŞekeroğlu Export Team"
        ),
    },
    "de": {
        "subject": "Vielen Dank für Ihren Besuch auf unserem Messestand",
        "body": (
            "Guten Tag,\n\n"
            "herzlichen Dank für das Gespräch an unserem Stand. Gerne senden wir Ihnen"
            " weitere Informationen zu unseren Produkten und prüfen gemeinsame"
            " Geschäftsmöglichkeiten.\n\n"
            "Teilen Sie uns bitte mit, wann wir Sie für ein kurzes Nachgespräch erreichen"
            " können.\n\n"
            "Mit freundlichen Grüßen\nŞekeroğlu Export Team"
        ),
    },
    "fr": {
        "subject": "Suite à notre rencontre sur le salon",
        "body": (
            "Bonjour,\n\n"
            "Nous vous remercions d’avoir visité notre stand lors du salon. Nous serions"
            " ravis de poursuivre nos échanges et de vous proposer des solutions adaptées"
            " à vos besoins.\n\n"
            "N’hésitez pas à nous indiquer vos disponibilités pour un échange complémentaire."
            "\n\n"
            "Cordialement,\nÉquipe Export Şekeroğlu"
        ),
    },
    "es": {
        "subject": "Seguimiento de nuestra reunión en la feria",
        "body": (
            "Hola,\n\n"
            "Muchas gracias por visitarnos durante la feria. Queremos continuar la"
            " conversación y presentarle nuestras soluciones ajustadas a sus necesidades."
            "\n\n"
            "Por favor, indíquenos cuándo podemos coordinar una reunión o llamada de"
            " seguimiento.\n\n"
            "Saludos cordiales,\nEquipo de Exportación Şekeroğlu"
        ),
    },
    "ar": {
        "subject": "متابعة بعد زيارتكم لجناحنا في المعرض",
        "body": (
            "مرحباً،\n\n"
            "نشكر لكم زيارتكم لجناحنا خلال المعرض واهتمامكم بمنتجات شكر أوغلو. يسعدنا"
            " تزويدكم بمعلومات إضافية وبحث فرص التعاون المشتركة.\n\n"
            "يرجى تزويدنا بالوقت المناسب للتواصل معكم والحديث عن الخطوات القادمة.\n\n"
            "مع أطيب التحيات،\nفريق تصدير شكر أوغلو"
        ),
    },
}

HOLIDAY_MAIL_TEMPLATES = {
    "Ramazan Bayramı": {
        "tr": {
            "subject": "Ramazan Bayramınız Kutlu Olsun",
            "body": (
                "<p>Değerli iş ortağımız,</p>"
                "<p>Ramazan Bayramınızı en içten dileklerimizle kutlarız. İş birliğimizin"
                " artarak devam etmesini diler, sevdiklerinizle birlikte sağlıklı ve mutlu"
                " günler temenni ederiz.</p>"
                "<p>Saygılarımızla,<br>Sekeroğlu Group</p>"
            ),
        },
        "en": {
            "subject": "Happy Eid al-Fitr",
            "body": (
                "<p>Dear business partner,</p>"
                "<p>We sincerely wish you a joyful and prosperous Eid al-Fitr. Thank you"
                " for your trust and cooperation.</p>"
                "<p>Best regards,<br>Sekeroğlu Group</p>"
            ),
        },
        "de": {
            "subject": "Frohes Zuckerfest",
            "body": (
                "<p>Sehr geehrter Geschäftspartner,</p>"
                "<p>wir wünschen Ihnen und Ihren Liebsten ein gesegnetes und"
                " erfolgreiches Zuckerfest. Vielen Dank für die vertrauensvolle"
                " Zusammenarbeit.</p>"
                "<p>Mit freundlichen Grüßen,<br>Sekeroğlu Group</p>"
            ),
        },
        "fr": {
            "subject": "Bonne fête de l’Aïd al-Fitr",
            "body": (
                "<p>Cher partenaire,</p>"
                "<p>Nous vous souhaitons une fête de l’Aïd al-Fitr pleine de joie et de"
                " prospérité. Merci pour votre confiance et votre collaboration.</p>"
                "<p>Cordialement,<br>Sekeroğlu Group</p>"
            ),
        },
        "es": {
            "subject": "Feliz Eid al-Fitr",
            "body": (
                "<p>Estimado socio,</p>"
                "<p>Le deseamos un Eid al-Fitr lleno de alegría y prosperidad. Gracias"
                " por su confianza y cooperación.</p>"
                "<p>Saludos cordiales,<br>Sekeroğlu Group</p>"
            ),
        },
        "ar": {
            "subject": "عيد فطر سعيد",
            "body": (
                "<p>شريكنا العزيز،</p>"
                "<p>نهنئكم بعيد الفطر المبارك ونتمنى لكم ولعائلتكم أياماً مليئة"
                " بالخير والنجاح. نشكركم على ثقتكم وشراكتكم المستمرة.</p>"
                "<p>مع أطيب التحيات،<br>مجموعة شكر أوغلو</p>"
            ),
        },
    },
    "Kurban Bayramı": {
        "tr": {
            "subject": "Kurban Bayramınız Kutlu Olsun",
            "body": (
                "<p>Değerli iş ortağımız,</p>"
                "<p>Kurban Bayramı'nın bereket ve mutluluk getirmesini diler, bugüne"
                " kadar gösterdiğiniz iş birliği için teşekkür ederiz.</p>"
                "<p>En iyi dileklerimizle,<br>Sekeroğlu Group</p>"
            ),
        },
        "en": {
            "subject": "Happy Eid al-Adha",
            "body": (
                "<p>Dear business partner,</p>"
                "<p>May this Eid al-Adha bring peace, happiness, and success to you and"
                " your loved ones. Thank you for your continued cooperation.</p>"
                "<p>Sincerely,<br>Sekeroğlu Group</p>"
            ),
        },
        "de": {
            "subject": "Frohes Opferfest",
            "body": (
                "<p>Sehr geehrter Geschäftspartner,</p>"
                "<p>möge das Opferfest Ihnen und Ihrem Team Frieden, Gesundheit und"
                " Erfolg bringen. Wir danken Ihnen für die gute Zusammenarbeit.</p>"
                "<p>Mit freundlichen Grüßen,<br>Sekeroğlu Group</p>"
            ),
        },
        "fr": {
            "subject": "Bonne fête de l’Aïd al-Adha",
            "body": (
                "<p>Cher partenaire,</p>"
                "<p>Que cette fête de l’Aïd al-Adha vous apporte paix, bonheur et"
                " réussite. Merci pour votre collaboration précieuse.</p>"
                "<p>Cordialement,<br>Sekeroğlu Group</p>"
            ),
        },
        "es": {
            "subject": "Feliz Eid al-Adha",
            "body": (
                "<p>Estimado socio,</p>"
                "<p>Que este Eid al-Adha le traiga paz, felicidad y éxito a usted y a"
                " su equipo. Gracias por su apoyo continuo.</p>"
                "<p>Saludos cordiales,<br>Sekeroğlu Group</p>"
            ),
        },
        "ar": {
            "subject": "عيد أضحى مبارك",
            "body": (
                "<p>شريكنا العزيز،</p>"
                "<p>نتمنى أن يحمل لكم عيد الأضحى المبارك السلام والسعادة والنجاح،"
                " ونشكركم على تعاونكم المستمر.</p>"
                "<p>مع خالص التقدير،<br>مجموعة شكر أوغلو</p>"
            ),
        },
    },
    "Yeni Yıl": {
        "tr": {
            "subject": "Mutlu Yıllar",
            "body": (
                "<p>Değerli iş ortağımız,</p>"
                "<p>Geride bıraktığımız yıl boyunca gösterdiğiniz destek için teşekkür"
                " ederiz. Yeni yılın size ve ekibinize sağlık, mutluluk ve başarı"
                " getirmesini dileriz.</p>"
                "<p>Sevgi ve saygılarımızla,<br>Sekeroğlu Group</p>"
            ),
        },
        "en": {
            "subject": "Happy New Year",
            "body": (
                "<p>Dear business partner,</p>"
                "<p>Thank you for the trust and partnership throughout the past year. We"
                " wish you and your team a healthy and prosperous New Year.</p>"
                "<p>Warm regards,<br>Sekeroğlu Group</p>"
            ),
        },
        "de": {
            "subject": "Frohes Neues Jahr",
            "body": (
                "<p>Sehr geehrter Geschäftspartner,</p>"
                "<p>vielen Dank für Ihre Unterstützung im vergangenen Jahr. Wir wünschen"
                " Ihnen und Ihrem Team ein gesundes und erfolgreiches neues Jahr.</p>"
                "<p>Mit besten Grüßen,<br>Sekeroğlu Group</p>"
            ),
        },
        "fr": {
            "subject": "Bonne année",
            "body": (
                "<p>Cher partenaire,</p>"
                "<p>Merci pour votre confiance tout au long de l’année écoulée. Nous vous"
                " souhaitons, à vous et à votre équipe, une nouvelle année pleine de"
                " santé et de réussite.</p>"
                "<p>Cordialement,<br>Sekeroğlu Group</p>"
            ),
        },
        "es": {
            "subject": "Feliz Año Nuevo",
            "body": (
                "<p>Estimado socio,</p>"
                "<p>Gracias por su confianza y colaboración durante el último año. Les"
                " deseamos a usted y a su equipo un Año Nuevo lleno de salud y"
                " prosperidad.</p>"
                "<p>Saludos cordiales,<br>Sekeroğlu Group</p>"
            ),
        },
        "ar": {
            "subject": "سنة جديدة سعيدة",
            "body": (
                "<p>شريكنا العزيز،</p>"
                "<p>نشكر لكم ثقتكم وشراكتكم طوال العام الماضي، ونتمنى لكم ولفريقكم"
                " عاماً جديداً مليئاً بالصحة والنجاح.</p>"
                "<p>مع أطيب الأمنيات،<br>مجموعة شكر أوغلو</p>"
            ),
        },
    },
}

HOLIDAY_FALLBACK_TEMPLATES = {
    "Ramazan Bayramı": {
        "subject": "Ramazan Bayramınız Kutlu Olsun / Happy Eid al-Fitr",
        "body": (
            "<p>Değerli iş ortağımız,</p>"
            "<p>Ramazan Bayramınızı en içten dileklerimizle kutlarız. İş birliğimizin"
            " artarak devam etmesini diler, sevdiklerinizle birlikte sağlıklı ve mutlu"
            " günler temenni ederiz.</p>"
            "<p>Saygılarımızla,<br>Sekeroğlu Group</p>"
            "<hr>"
            "<p>Dear business partner,</p>"
            "<p>We sincerely wish you a joyful and prosperous Eid al-Fitr. Thank you for"
            " your trust and cooperation.</p>"
            "<p>Best regards,<br>Sekeroğlu Group</p>"
        ),
    },
    "Kurban Bayramı": {
        "subject": "Kurban Bayramınız Kutlu Olsun / Happy Eid al-Adha",
        "body": (
            "<p>Değerli iş ortağımız,</p>"
            "<p>Kurban Bayramı'nın bereket ve mutluluk getirmesini diler, bugüne kadar"
            " gösterdiğiniz iş birliği için teşekkür ederiz.</p>"
            "<p>En iyi dileklerimizle,<br>Sekeroğlu Group</p>"
            "<hr>"
            "<p>Dear business partner,</p>"
            "<p>May this Eid al-Adha bring peace, happiness and success to you and your"
            " loved ones.</p>"
            "<p>Sincerely,<br>Sekeroğlu Group</p>"
        ),
    },
    "Yeni Yıl": {
        "subject": "Mutlu Yıllar / Happy New Year",
        "body": (
            "<p>Değerli iş ortağımız,</p>"
            "<p>Geride bıraktığımız yıl boyunca gösterdiğiniz destek için teşekkür ederiz."
            " Yeni yılın size ve ekibinize sağlık, mutluluk ve başarı getirmesini dileriz.</p>"
            "<p>Sevgi ve saygılarımızla,<br>Sekeroğlu Group</p>"
            "<hr>"
            "<p>Dear business partner,</p>"
            "<p>Thank you for the trust and partnership throughout the past year. Wishing"
            " you a healthy and prosperous New Year.</p>"
            "<p>Warm regards,<br>Sekeroğlu Group</p>"
        ),
    },
}

LANGUAGE_LABELS = {
    "tr": "Türkçe",
    "en": "English",
    "de": "Deutsch",
    "fr": "Français",
    "es": "Español",
    "ar": "العربية",
}


def get_holiday_template_content(template_name, language_code):
    template = HOLIDAY_MAIL_TEMPLATES.get(template_name, {})
    content = template.get(language_code)
    if content:
        return content, False
    fallback = HOLIDAY_FALLBACK_TEMPLATES.get(template_name)
    if fallback:
        return fallback, True
    return None, False



# --- LOGO (WEB LINKİNDEN AL) ---
logo_url = "https://www.sekeroglugroup.com/storage/settings/xdp5r6DZIFJMNGOStqwvKCiVHDhYxA84jFr61TNp.svg"

col1, col2 = st.columns([3, 7])
with col1:
    st.image(logo_url, width=300)
with col2:
    st.markdown("""
        <style>
        .block-container { padding-top: 0.2rem !important; }
        </style>
        <div style="display:flex; flex-direction:column; align-items:flex-start; width:100%; margin-bottom:10px;">
            <h1 style="color: #219A41; font-weight: bold; font-size: 2.8em; letter-spacing:2px; margin:0; margin-top:-8px;">
                ŞEKEROĞLU İHRACAT CRM
            </h1>
        </div>
    """, unsafe_allow_html=True)



@st.cache_resource
def get_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    return GoogleDrive(gauth)
drive = get_drive()

downloaded = drive.CreateFile({'id': EXCEL_FILE_ID})
downloaded.FetchMetadata(fetch_all=True)
downloaded.GetContentFile("temp.xlsx")


def load_dataframes_from_excel(path: str = "temp.xlsx"):
    global df_musteri, df_kayit, df_teklif, df_proforma, df_evrak, df_eta, df_fuar_musteri

    if os.path.exists(path):
        try:
            df_musteri = pd.read_excel(path, sheet_name=0)
        except Exception:
            df_musteri = pd.DataFrame(columns=[
                "Müşteri Adı", "Telefon", "E-posta", "Adres", "Ülke", "Satış Temsilcisi", "Kategori", "Durum", "Vade (Gün)", "Ödeme Şekli"
            ])
        try:
            df_kayit = pd.read_excel(path, sheet_name="Kayıtlar")
        except Exception:
            df_kayit = pd.DataFrame(columns=["Müşteri Adı", "Tarih", "Tip", "Açıklama"])
        try:
            df_teklif = pd.read_excel(path, sheet_name="Teklifler")
        except Exception:
            df_teklif = pd.DataFrame(columns=[
                "Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Ürün/Hizmet", "Açıklama", "Durum", "PDF"
            ])
        try:
            df_proforma = pd.read_excel(path, sheet_name="Proformalar")
            for col in ["Proforma No", "Vade", "Sevk Durumu"]:
                if col not in df_proforma.columns:
                    df_proforma[col] = ""
        except Exception:
            df_proforma = pd.DataFrame(columns=[
                "Müşteri Adı", "Tarih", "Proforma No", "Tutar", "Açıklama", "Durum", "PDF", "Sipariş Formu", "Vade", "Sevk Durumu"
            ])
        try:
            df_evrak = pd.read_excel(path, sheet_name="Evraklar")
            for col in ["Yük Resimleri", "EK Belgeler"]:
                if col not in df_evrak.columns:
                    df_evrak[col] = ""
        except Exception:
            df_evrak = pd.DataFrame(columns=[
                "Müşteri Adı", "Fatura No", "Fatura Tarihi", "Vade Tarihi", "Tutar",
                "Commercial Invoice", "Sağlık Sertifikası", "Packing List",
                "Konşimento", "İhracat Beyannamesi", "Fatura PDF", "Sipariş Formu",
                "Yük Resimleri", "EK Belgeler"
            ])
        try:
            df_eta = pd.read_excel(path, sheet_name="ETA")
        except Exception:
            for col in ETA_COLUMNS:
                if col not in df_eta.columns:
                    df_eta[col] = ""
            extra_cols = [col for col in df_eta.columns if col not in ETA_COLUMNS]
            df_eta = df_eta.reindex(columns=ETA_COLUMNS + extra_cols, fill_value="")            
        try:
            df_fuar_musteri = pd.read_excel(path, sheet_name="FuarMusteri")
        except Exception:
            df_fuar_musteri = pd.DataFrame(columns=[
                "Fuar Adı", "Müşteri Adı", "Ülke", "Telefon", "E-mail", "Açıklamalar", "Tarih"
            ])
    else:
        df_musteri = pd.DataFrame(columns=[
            "Müşteri Adı", "Telefon", "E-posta", "Adres", "Ülke", "Satış Temsilcisi", "Kategori", "Durum", "Vade (Gün)", "Ödeme Şekli"
        ])
        df_kayit = pd.DataFrame(columns=["Müşteri Adı", "Tarih", "Tip", "Açıklama"])
        df_teklif = pd.DataFrame(columns=[
            "Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Ürün/Hizmet", "Açıklama", "Durum", "PDF"
        ])
        df_proforma = pd.DataFrame(columns=[
            "Müşteri Adı", "Tarih", "Proforma No", "Tutar", "Açıklama", "Durum", "PDF", "Sipariş Formu", "Vade", "Sevk Durumu"
        ])
        df_evrak = pd.DataFrame(columns=[
            "Müşteri Adı", "Fatura No", "Fatura Tarihi", "Vade Tarihi", "Tutar",
            "Ödenen Tutar", "Commercial Invoice", "Sağlık Sertifikası", "Packing List",
            "Konşimento", "İhracat Beyannamesi", "Fatura PDF", "Sipariş Formu",
            "Yük Resimleri", "EK Belgeler"
        ])
        df_eta = pd.DataFrame(columns=ETA_COLUMNS)        
        df_fuar_musteri = pd.DataFrame(columns=[
            "Fuar Adı", "Müşteri Adı", "Ülke", "Telefon", "E-mail", "Açıklamalar", "Tarih"
        ])


load_dataframes_from_excel()

def update_excel():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_musteri.to_excel(writer, sheet_name="Sayfa1", index=False)
        df_kayit.to_excel(writer, sheet_name="Kayıtlar", index=False)
        df_teklif.to_excel(writer, sheet_name="Teklifler", index=False)
        df_proforma.to_excel(writer, sheet_name="Proformalar", index=False)
        df_evrak.to_excel(writer, sheet_name="Evraklar", index=False)
        df_eta.to_excel(writer, sheet_name="ETA", index=False)
        df_fuar_musteri.to_excel(writer, sheet_name="FuarMusteri", index=False)
    buffer.seek(0)
    with open("temp.xlsx", "wb") as f:
        f.write(buffer.read())
    downloaded.SetContentFile("temp.xlsx")
    downloaded.Upload()



def sync_excel_bidirectional():
    global downloaded

    try:
        downloaded.FetchMetadata(fetch_all=True)
        remote_raw = downloaded.get('modifiedDate') or downloaded.get('modifiedTime')
        remote_ts = None
        if remote_raw:
            remote_ts = pd.to_datetime(remote_raw, utc=True).tz_convert(None).to_pydatetime()
    except Exception as e:
        st.session_state.sync_status = ("error", f"Drive meta verisi alınamadı: {e}")
        return

    local_exists = os.path.exists("temp.xlsx")
    local_ts = None
    if local_exists:
        try:
            local_ts = datetime.datetime.fromtimestamp(os.path.getmtime("temp.xlsx"))
        except Exception:
            local_ts = None

    tolerance = datetime.timedelta(seconds=2)

    if not local_exists or (remote_ts and (local_ts is None or remote_ts - local_ts > tolerance)):
        try:
            downloaded.GetContentFile("temp.xlsx")
            load_dataframes_from_excel()
            st.session_state.sync_status = ("success", "Google Drive dosyası daha güncel bulundu; yerel kopya yenilendi.")
        except Exception as e:
            st.session_state.sync_status = ("error", f"Drive'dan dosya indirilirken hata oluştu: {e}")
        return

    if remote_ts and local_ts and (local_ts - remote_ts > tolerance):
        try:
            downloaded.SetContentFile("temp.xlsx")
            downloaded.Upload()
            st.session_state.sync_status = ("success", "Yerel dosya daha güncel bulundu; Drive üzerindeki dosya güncellendi.")
        except Exception as e:
            st.session_state.sync_status = ("error", f"Drive'a dosya yüklenirken hata oluştu: {e}")
    else:
        st.session_state.sync_status = ("info", "Dosyalar zaten senkron görünüyor.")


if st.session_state.pop("_sync_requested", False):
    sync_excel_bidirectional()


# ===========================
# ==== GOOGLE SHEETS (MÜŞTERİ) SENKRON
# ===========================
def _df_to_values(df: pd.DataFrame):
    def _cell(v):
        if pd.isna(v): return ""
        if isinstance(v, (pd.Timestamp, datetime.date, datetime.datetime)):
            try: return pd.to_datetime(v).date().isoformat()
            except: return str(v)
        if isinstance(v, (np.bool_, bool)): return bool(v)
        return v
    header = list(df.columns)
    rows = df.applymap(_cell).values.tolist()
    return [header] + rows

def write_customers_to_gsheet(df_customers: pd.DataFrame) -> bool:
    if "sheets_svc" not in globals() or sheets_svc is None:
        st.error("Sheets servisi hazır değil!")
        return False
    try:
        if df_customers is None or df_customers.empty:
            st.warning("Müşteri tablosu boş, Sheets’e yazılacak bir şey yok.")
            return False
        sheet = sheets_svc.spreadsheets()
        # clear
        execute_with_retry(sheet.values().batchClear(
            spreadsheetId=SHEET_ID,
            body={"ranges":[f"{MUSTERI_SHEET_NAME}!A:ZZ"]}
        ))
        # write
        values = _df_to_values(df_customers)
        execute_with_retry(sheet.values().update(
            spreadsheetId=SHEET_ID,
            range=f"{MUSTERI_SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": values}
        ))
        st.info(f"{MUSTERI_SHEET_NAME} sayfasına {len(df_customers)} satır yazıldı.")
        return True
    except Exception as e:
        st.error(f"Sheets yazma hatası: {e}")
        return False

def push_customers_throttled():
    now = datetime.datetime.utcnow().timestamp()
    last = st.session_state.get("_last_sheet_write_ts", 0)
    if now - last < 10:  # 10 sn içinde tekrar yazma (429 riski azalt)
        return False
    ok = write_customers_to_gsheet(df_musteri)
    if ok:
        st.session_state["_last_sheet_write_ts"] = now
    return ok

import smtplib
from email.message import EmailMessage

# Yeni cari için txt dosyasını oluşturma fonksiyonu
def yeni_cari_txt_olustur(cari_dict, file_path="yeni_cari.txt"):
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(
            f"Müşteri Adı: {cari_dict['Müşteri Adı']}\n"
            f"Telefon: {cari_dict['Telefon']}\n"
            f"E-posta: {cari_dict['E-posta']}\n"
            f"Adres: {cari_dict['Adres']}\n"
            f"Ülke: {cari_dict.get('Ülke', '')}\n"
            f"Satış Temsilcisi: {cari_dict.get('Satış Temsilcisi', '')}\n"
            f"Kategori: {cari_dict.get('Kategori', '')}\n"
            f"Durum: {cari_dict.get('Durum', '')}\n"
            f"Vade (Gün): {cari_dict.get('Vade (Gün)', '')}\n"
            f"Ödeme Şekli: {cari_dict.get('Ödeme Şekli', '')}\n"
            f"Para Birimi: {cari_dict.get('Para Birimi', '')}\n"  # Para birimini de ekliyoruz
            f"DT Seçimi: {cari_dict.get('DT Seçimi', '')}\n"  # DT seçimini de ekliyoruz
        )
# --- E-POSTA İMZA YARDIMCILARI ---

SIGNATURE_PROFILES = {
    "admin": {
        "full_name": "KEMAL İLKER ÇELİKKALKAN",
        "title": "Export Manager",
    },
    "export1": {
        "full_name": "HÜSEYİN POLAT",
        "title": "Export Area Sales Manager",
    },
    "Boss": {
        "full_name": "FERHAT ŞEKEROĞLU",
        "title": "MEMBER OF BOARD",
    },
}

SIGNATURE_BASE_INFO = {
    "company": "ŞEKEROĞLU GROUP",
    "department": "International Sales & Export",
    "phone": "+90 (342) 337 09 09",
    "email": "export1@sekeroglugroup.com",
    "website": "https://www.sekeroglugroup.com",
    "address": "Sanayi mah. 60129 No'lu Cad. No : 7 Şehitkamil / Gaziantep",
    }


def _active_signature_info():
    user = st.session_state.get("user") if "user" in st.session_state else None
    info = SIGNATURE_PROFILES.get(user)
    if not info:
        info = {
            "full_name": "ŞEKEROĞLU EXPORT TEAM",
            "title": "International Sales Representative",
        }
    return info


def text_signature() -> str:
    info = _active_signature_info()
    base = SIGNATURE_BASE_INFO
    lines = [
        info["full_name"],
        info["title"],
        f"{base['company']} | {base['department']}",
        f"Telefon: {base['phone']}",
        f"E-posta: {base['email']}",
        f"Web: {base['website']}",
        f"Adres: {base['address']}",
    ]
    return "\n".join(lines)


def html_signature() -> str:
    info = _active_signature_info()
    base = SIGNATURE_BASE_INFO
    name = html.escape(info["full_name"])
    title = html.escape(info["title"])
    company = html.escape(base["company"])
    department = html.escape(base["department"])
    phone = html.escape(base["phone"])
    email = html.escape(base["email"])
    website = html.escape(base["website"])
    address = html.escape(base["address"])

    return (
        "<table cellpadding=\"0\" cellspacing=\"0\" style=\"font-family:Arial,sans-serif;color:#333333;\">"
        f"<tr><td style='font-size:16px;font-weight:bold;color:#219A41;'>{name}</td></tr>"
        f"<tr><td style='font-size:13px;color:#555555;'>{title}</td></tr>"
        "<tr><td style='padding-top:10px;font-size:13px;color:#333333;'>"
        f"<div style='font-weight:bold;color:#219A41;'>{company}</div>"
        f"<div>{department}</div>"
        f"<div>Telefon: {phone}</div>"
        f"<div>E-posta: <a href='mailto:{email}' style='color:#219A41;text-decoration:none;'>{email}</a></div>"
        f"<div>Web: <a href='{website}' style='color:#219A41;text-decoration:none;'>{website}</a></div>"
        f"<div>{address}</div>"
        "</td></tr>"
        "</table>"
    )


# E-posta göndermek için fonksiyon
def send_email(to_email, subject, body, attachments=None, fallback_txt_path=None):
    from_email = "todo@sekeroglugroup.com"  # Gönderen e-posta adresi
    password = "vbgvforwwbcpzhxf"  # Gönderen e-posta şifresi

    # E-posta mesajını oluştur
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email

    
    if isinstance(to_email, (str, bytes)):
        recipients = [to_email.decode() if isinstance(to_email, bytes) else to_email]
    else:
        recipients = [addr for addr in to_email if addr]

    if not recipients:
        raise ValueError("En az bir geçerli alıcı e-posta adresi sağlanmalıdır.")

    msg["To"] = from_email
    msg["Bcc"] = ", ".join(recipients)
    msg.set_content(body)

    prepared_attachments = attachments or []

    for attachment in prepared_attachments:
        if not isinstance(attachment, (tuple, list)) or len(attachment) != 3:
            continue

        filename, content_bytes, mime_type = attachment

        if isinstance(content_bytes, memoryview):
            content_bytes = content_bytes.tobytes()

        if not isinstance(content_bytes, (bytes, bytearray)):
            content_bytes = bytes(content_bytes)

        mime = (mime_type or "application/octet-stream").strip()
        if "/" in mime:
            maintype, subtype = mime.split("/", 1)
        else:
            maintype, subtype = mime, "octet-stream"

        msg.add_attachment(
           bytes(content_bytes),
            maintype=maintype,
            subtype=subtype,
            filename=filename,
        )

    if not prepared_attachments and fallback_txt_path:
        try:
            with open(fallback_txt_path, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype="text",
                    subtype="plain",
                    filename=os.path.basename(fallback_txt_path) or "yeni_cari.txt",
                )
        except FileNotFoundError:
            pass


    # E-posta göndermek için SMTP kullan
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(from_email, password)
        smtp.send_message(msg)

def extract_unique_emails(email_series: pd.Series) -> list:
    emails = []
    if email_series is None:
        return emails
    for raw in email_series.dropna():
        raw_str = str(raw).strip()
        if not raw_str:
            continue
        parts = re.split(r"[;,\s]+", raw_str)
        for part in parts:
            address = part.strip()
            if address:
                emails.append(address)
    # Benzersiz ve alfabetik sırada döndür
    seen = {}
    for mail in emails:
        key = mail.lower()
        if key not in seen:
            seen[key] = mail
    return sorted(seen.values(), key=lambda x: x.lower())


def send_fair_bulk_email(to_emails, subject, body, attachments=None, embed_images=None, inline_cid_map=None):
    if not to_emails:
        raise ValueError("E-posta alıcı listesi boş olamaz.")

    from_email = "todo@sekeroglugroup.com"
    password = "vbgvforwwbcpzhxf"

    embed_images = EMBED_IMAGES if embed_images is None else bool(embed_images)
    inline_cid_map = inline_cid_map or {}

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = from_email
    msg["Bcc"] = ", ".join(to_emails)
    body_text = (body or "").strip()
    if body_text:
        plain_body = f"{body_text}\n\n{text_signature()}"
    else:
        plain_body = text_signature()
    msg.set_content(plain_body)

    escaped_body = html.escape(body_text).replace("\n", "<br>") if body_text else ""

    attachments = attachments or []
    inline_images = []
    download_errors = []
    other_attachments = []
    
    for uploaded_file in attachments:
        if uploaded_file is None:
            continue
        try:
            file_bytes = uploaded_file.getvalue()
        except Exception as exc:
            download_errors.append(f"{getattr(uploaded_file, 'name', 'Dosya')} okunamadı: {exc}")
            continue

        maintype, subtype = "application", "octet-stream"
        mime_type = uploaded_file.type or ""
        if mime_type:
            try:
                maintype, subtype = mime_type.split("/", 1)
            except ValueError:
                maintype, subtype = mime_type, "octet-stream"

        attachment_key = f"{uploaded_file.name}:{len(file_bytes)}"

        if embed_images and maintype == "image":
            cid = make_msgid()
            cid = inline_cid_map.get(attachment_key, cid)
            inline_images.append({
                "data": file_bytes,
                "maintype": maintype,
                "subtype": subtype,
                "filename": uploaded_file.name,
                "cid": cid,
                "key": attachment_key,
            })
        else:
            other_attachments.append({
                "data": file_bytes,
                "maintype": maintype,
                "subtype": subtype,
                "filename": uploaded_file.name,
            })

    if download_errors:
        raise RuntimeError("; ".join(download_errors))

    image_blocks = []
    if embed_images and inline_images:
        for image in inline_images:
            cid_clean = image["cid"].strip("<>")
            alt_text = html.escape(image["filename"]) if image["filename"] else "Görsel"
            image_blocks.append(
                "<div style='margin-top:12px; text-align:center;'>"
                f"<img src=\"cid:{cid_clean}\" alt=\"{alt_text}\" style='max-width:100%; height:auto;'>"
                "</div>"
            )

    html_sections = ["<div style='font-family:Arial,sans-serif;color:#333333;'>"]
    if escaped_body:
        html_sections.append(escaped_body)
    if image_blocks:
        html_sections.extend(image_blocks)
    html_sections.append("</div>")
    html_sections.append(f"<div style='margin-top:16px;'>{html_signature()}</div>")
    html_body = "".join(html_sections)

    html_part = msg.add_alternative(html_body, subtype="html")

    for image in inline_images:
        html_part.add_related(
            image["data"],
            maintype=image["maintype"],
            subtype=image["subtype"],
            cid=image["cid"],
            filename=image["filename"] or None,
        )

    for attachment in other_attachments:
        msg.add_attachment(
            attachment["data"],
            maintype=attachment["maintype"],
            subtype=attachment["subtype"],
            filename=attachment["filename"],
        )

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(from_email, password)
            smtp.send_message(msg)
    except smtplib.SMTPException as exc:
        raise RuntimeError(f"SMTP hatası: {exc}") from exc





# ========= /ŞIK MENÜ =========
# ===========================
# ==== ŞIK SIDEBAR MENÜ
# ===========================

menuler = [
    ("Özet Ekran", "📊"),
    ("Cari Kayıtlar", "🧑‍💼"),
    ("Müşteri Portföyü", "📒"),
    ("Etkileşim Günlüğü", "☎️"),
    ("Teklif Yönetimi", "💰"),
    ("Proforma Yönetimi", "📄"),
    ("Sipariş Operasyonları", "🚚"),
    ("Fatura işlemleri", "📑"),
    ("Tahsilat Planı", "⏰"),
    ("ETA İzleme", "🛳️"),
    ("Fuar Kayıtları", "🎫"),
    ("İçerik Arşivi", "🗂️"),
    ("Satış Analitiği", "📈"),
    ("Özel Gün Tebrikleri", "🎉"),
    ("Help & Support", "❓"),
    ("Settings", "⚙️"),
]

# 2) Tüm kullanıcılar için aynı menüler
USER_MENU_PERMISSIONS = {
    "Muhammed": {"ETA İzleme"},
    "Muhammed": {"ETA İzleme", "Fatura işlemleri"},    
}


def resolve_allowed_menus(username):
    allowed_names = USER_MENU_PERMISSIONS.get(username)
    if not allowed_names:
        return menuler

    filtered = [item for item in menuler if item[0] in allowed_names]
    return filtered if filtered else menuler


# 2) Kullanıcıya göre menüleri sınırla
allowed_menus = resolve_allowed_menus(st.session_state.user)

# 3) Etiketler ve haritalar
labels = [isim for (isim, _ikon) in allowed_menus]
name_by_label = {isim: isim for (isim, _ikon) in allowed_menus}
label_by_name = {isim: isim for (isim, _ikon) in allowed_menus}

# 4) Varsayılan state
if "menu_state" not in st.session_state:
    st.session_state.menu_state = allowed_menus[0][0]

# 5) CSS (kart görünümü; input gizlenmiyor—erişilebilir kalır)
st.sidebar.markdown(
    """
    <style>
    div[data-testid="stSidebar"] .stRadio > div {
        gap: 4px !important;
    }

    div[data-testid="stSidebar"] .stRadio label {
        border-radius: 12px;
        padding: 11px 16px;
        margin-bottom: 2px;
        border: 1px solid transparent;
        display: flex;
        align-items: center;
        background: transparent;
        transition: border-color 0.2s ease, background 0.2s ease, color 0.2s ease, transform 0.12s ease;
    }

    div[data-testid="stSidebar"] .stRadio label:hover {
        border-color: rgba(148, 163, 184, 0.28);
        background: rgba(148, 163, 184, 0.08);
        transform: translateX(2px);
    }

    div[data-testid="stSidebar"] .stRadio label span {
        font-weight: 500;
        color: #94a3b8;
        font-size: 0.95rem;
        letter-spacing: 0.01em;
    }

    div[data-testid="stSidebar"] .stRadio [aria-checked="true"] {
        background: rgba(56, 189, 248, 0.16);
        border-color: rgba(56, 189, 248, 0.55);
        box-shadow: inset 3px 0 0 #38bdf8;
    }

    div[data-testid="stSidebar"] .stRadio [aria-checked="true"] span {
        color: #f8fafc;
        font-weight: 600;
    }

    div[data-testid="stSidebar"] .streamlit-expanderHeader p {
        color: #94a3b8 !important;
        font-weight: 500;
        letter-spacing: 0.02em;
    }

    div[data-testid="stSidebar"] .streamlit-expanderHeader:hover {
        color: #38bdf8 !important;
    }

    div[data-testid="stSidebar"] .streamlit-expanderContent {
        background: rgba(15, 23, 42, 0.65);
        border-left: 1px solid rgba(148, 163, 184, 0.18);
        padding-left: 0.75rem;
    }

    div[data-testid="stSidebar"] .stText, div[data-testid="stSidebar"] .stMarkdown {
        color: #94a3b8 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 6) Callback: seçilince anında state yaz (tek tıkta geçiş)
def _on_menu_change():
    sel_label = st.session_state.menu_radio_label
    st.session_state.menu_state = name_by_label.get(sel_label, allowed_menus[0][0])

# 7) Radio’yu mevcut state’e göre başlat
current_label = label_by_name.get(st.session_state.menu_state, labels[0])
current_index = labels.index(current_label) if current_label in labels else 0

st.sidebar.radio(
    "Menü",
    labels,
    index=current_index,
    label_visibility="collapsed",
    key="menu_radio_label",
    on_change=_on_menu_change
)

# 8) Kullanım: seçili menü adı
menu = st.session_state.menu_state
allowed_menu_names = {isim for (isim, _ikon) in allowed_menus}
if menu not in allowed_menu_names:
    menu = allowed_menus[0][0]
    st.session_state.menu_state = menu


# Sidebar: manuel senkron
with st.sidebar.expander("🔄 Sheets Senkron"):
    if st.button("Müşterileri Sheets’e Yaz"):
        push_customers_throttled()



### ===========================
### === GENEL BAKIŞ (Vade Durumu Dahil) ===
### ===========================

if menu == "Özet Ekran":
    st.markdown(
        "<h2 style='color:#38bdf8; font-weight:bold;'>ŞEKEROĞLU İHRACAT CRM - Özet Ekran</h2>",
        unsafe_allow_html=True,
    )

    invoices_df = df_evrak.copy()
    if "Tutar" not in invoices_df.columns:
        invoices_df["Tutar"] = 0
    invoices_df["Tutar_num"] = invoices_df["Tutar"].apply(smart_to_num).fillna(0.0)
    toplam_fatura_tutar = float(invoices_df["Tutar_num"].sum())

    date_col = next((col for col in ["Fatura Tarihi", "Tarih"] if col in invoices_df.columns), None)
    today_norm = pd.Timestamp.today().normalize()

    if date_col:
        invoices_df[date_col] = pd.to_datetime(invoices_df[date_col], errors="coerce")
    # ---- Bekleyen Teklifler ----
    st.markdown("### Bekleyen Teklifler")
    bekleyen_teklifler = df_teklif[df_teklif["Durum"] == "Açık"] if "Durum" in df_teklif.columns else pd.DataFrame()
    try:
        toplam_teklif = pd.to_numeric(bekleyen_teklifler["Tutar"], errors="coerce").sum()
    except:
        toplam_teklif = 0
    st.markdown(f"<div style='font-size:1.3em; color:#11998e; font-weight:bold;'>Toplam: {toplam_teklif:,.2f} USD</div>", unsafe_allow_html=True)
    if bekleyen_teklifler.empty:
        st.info("Bekleyen teklif yok.")
    else:
        st.dataframe(bekleyen_teklifler[["Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Ürün/Hizmet", "Açıklama"]], use_container_width=True)

    # ---- Bekleyen Proformalar ----
    st.markdown("### Bekleyen Proformalar")
    bekleyen_proformalar = df_proforma[df_proforma["Durum"] == "Beklemede"] if "Durum" in df_proforma.columns else pd.DataFrame()
    try:
        toplam_proforma = pd.to_numeric(bekleyen_proformalar["Tutar"], errors="coerce").sum()
    except:
        toplam_proforma = 0
    st.markdown(f"<div style='font-size:1.3em; color:#f7971e; font-weight:bold;'>Toplam: {toplam_proforma:,.2f} USD</div>", unsafe_allow_html=True)
    if bekleyen_proformalar.empty:
        st.info("Bekleyen proforma yok.")
    else:
        st.dataframe(bekleyen_proformalar[["Müşteri Adı", "Proforma No", "Tarih", "Tutar", "Vade (gün)", "Açıklama"]], use_container_width=True)

    # ---- Sevk Bekleyen Siparişler ----

    if "Sevk Durumu" not in df_proforma.columns:
        df_proforma["Sevk Durumu"] = ""
    if "Ülke" not in df_proforma.columns:
        df_proforma["Ülke"] = ""
    if "Termin Tarihi" not in df_proforma.columns:
        df_proforma["Termin Tarihi"] = ""     
    sevk_bekleyenler = df_proforma[(df_proforma["Durum"] == "Siparişe Dönüştü") & (~df_proforma["Sevk Durumu"].isin(["Sevkedildi", "Ulaşıldı"]))] if "Durum" in df_proforma.columns else pd.DataFrame()
    sevk_bekleyen_sayisi = len(sevk_bekleyenler)
    st.markdown(f"### Sevk Bekleyen Siparişler ({sevk_bekleyen_sayisi} Adet)")
    try:
        toplam_siparis = pd.to_numeric(sevk_bekleyenler["Tutar"], errors="coerce").sum()
    except:
        toplam_siparis = 0
    st.markdown(f"<div style='font-size:1.3em; color:#185a9d; font-weight:bold;'>Toplam: {toplam_siparis:,.2f} USD</div>", unsafe_allow_html=True)
    if sevk_bekleyenler.empty:
        st.info("Sevk bekleyen sipariş yok.")
    else:
        sevk_bekleyenler = sevk_bekleyenler.copy()
        sevk_bekleyenler["Termin Tarihi"] = pd.to_datetime(
            sevk_bekleyenler["Termin Tarihi"], errors="coerce"
        )
        sevk_bekleyenler["Proforma Tarihi"] = pd.to_datetime(
            sevk_bekleyenler["Tarih"], errors="coerce"
        )        
        sevk_bekleyenler = sevk_bekleyenler.sort_values(
            by="Termin Tarihi", ascending=True, na_position="last"
        )

        display_df = sevk_bekleyenler.copy()
        display_df["Termin - Bugün Farkı (Gün)"] = (
            display_df["Termin Tarihi"] - today_norm
        ).dt.days
        display_df["Sipariş Üzerinden Geçen Gün"] = (
            display_df["Termin Tarihi"] - display_df["Proforma Tarihi"]
        ).dt.days

        for tarih_kolon in ["Termin Tarihi", "Proforma Tarihi"]:
            display_df[tarih_kolon] = (
                display_df[tarih_kolon]
                .dt.strftime("%d/%m/%Y")
                .fillna("")
                .replace({"NaT": ""})
            )

        for gun_kolon in ["Termin - Bugün Farkı (Gün)", "Sipariş Üzerinden Geçen Gün"]:
            display_df[gun_kolon] = display_df[gun_kolon].apply(
                lambda x: "" if pd.isna(x) else int(x)
            )

        st.dataframe(
            display_df[
                [
                    "Müşteri Adı",
                    "Ülke",
                    "Proforma No",
                    "Proforma Tarihi",
                    "Termin Tarihi",
                    "Termin - Bugün Farkı (Gün)",
                    "Sipariş Üzerinden Geçen Gün",                   
                    "Tutar",
                    "Açıklama",
                ]
            ],
            use_container_width=True,
        )

    # ---- Yolda Olan Siparişler ----
    st.markdown("### ETA Takibindeki Siparişler")
    eta_yolda = df_proforma[(df_proforma["Sevk Durumu"] == "Sevkedildi") & (~df_proforma["Sevk Durumu"].isin(["Ulaşıldı"]))] if "Sevk Durumu" in df_proforma.columns else pd.DataFrame()
    sevkiyat_sayisi = len(eta_yolda)
    st.markdown(
        f"<div style='font-size:1.3em; color:#c471f5; font-weight:bold;'>Sevkiyat Sayısı: {sevkiyat_sayisi}</div>",
        unsafe_allow_html=True,
    )
    if eta_yolda.empty:
        st.info("Yolda olan (sevk edilmiş) sipariş yok.")
    else:
        eta_display = eta_yolda.copy()
        eta_display["Proforma No"] = eta_display["Proforma No"].astype(str)

        if not df_eta.empty:
            eta_lookup = df_eta.copy()
            eta_lookup["Proforma No"] = eta_lookup["Proforma No"].astype(str)
            eta_lookup = eta_lookup.sort_values(by="ETA Tarihi", kind="stable")
            eta_lookup = eta_lookup.drop_duplicates(subset=["Proforma No"], keep="last")
            eta_display = eta_display.merge(
                eta_lookup[["Proforma No", "ETA Tarihi"]],
                on="Proforma No",
                how="left",
            )
        else:
            eta_display["ETA Tarihi"] = pd.NaT

        eta_display["ETA Tarihi"] = pd.to_datetime(eta_display["ETA Tarihi"], errors="coerce")
        eta_display["Kalan Gün"] = (
            eta_display["ETA Tarihi"] - today_norm
        ).dt.days
        eta_display["Kalan Gün"] = eta_display["Kalan Gün"].apply(
            lambda x: "" if pd.isna(x) else int(x)
        )        
        eta_display = eta_display.sort_values(
            by="ETA Tarihi",
            ascending=True,
            na_position="last",
            kind="stable",
        )
        eta_display["ETA Tarihi"] = eta_display["ETA Tarihi"].dt.strftime("%d/%m/%Y")
        eta_display["ETA Tarihi"] = eta_display["ETA Tarihi"].fillna("").replace({"NaT": ""})

        st.dataframe(
            eta_display[[
                "Müşteri Adı",
                "Ülke",
                "Proforma No",
                "Tarih",
                "ETA Tarihi",
                "Tutar",
                "Kalan Gün",
                "Açıklama",
            ]],
            use_container_width=True,
        )
    # ---- Son Teslim Edilen Siparişler ----
    st.markdown("### Son Teslim Edilen 5 Sipariş")
    if "Sevk Durumu" in df_proforma.columns:
        teslim_edilenler = df_proforma[df_proforma["Sevk Durumu"] == "Ulaşıldı"].copy()
        if not teslim_edilenler.empty:
            # Ana kolonları normalize et
            for text_col in ["Proforma No", "Müşteri Adı", "Ülke"]:
                if text_col in teslim_edilenler.columns:
                    teslim_edilenler[text_col] = (
                        teslim_edilenler[text_col]
                        .astype(str)
                        .str.strip()
                    )

            # Tarih kolonlarını datetime'a çevir
            date_cols = [
                "Proforma Tarihi",
                "Tarih",
                "Termin Tarihi",
                "Sevk Tarihi",
                "Ulaşma Tarihi",
            ]
            for tarih_kolon in date_cols:
                if tarih_kolon in teslim_edilenler.columns:
                    teslim_edilenler[tarih_kolon] = pd.to_datetime(
                        teslim_edilenler[tarih_kolon], errors="coerce"
                    )

            if "Proforma Tarihi" not in teslim_edilenler.columns and "Tarih" in teslim_edilenler.columns:
                teslim_edilenler["Proforma Tarihi"] = teslim_edilenler["Tarih"]

            # Son teslim edilenleri sıralayıp tekrarlayan proformaları ele
            sort_candidates = ["Ulaşma Tarihi", "Sevk Tarihi", "Proforma Tarihi"]
            sort_col = next(
                (col for col in sort_candidates if col in teslim_edilenler.columns),
                None,
            )

            if sort_col is not None:
                teslim_edilenler = teslim_edilenler.sort_values(
                    by=sort_col, ascending=False, na_position="last"
                )

            if "Proforma No" in teslim_edilenler.columns:
                teslim_edilenler = teslim_edilenler.drop_duplicates(
                    subset=["Proforma No"], keep="first"
                )

            teslim_edilenler = teslim_edilenler.head(5).reset_index(drop=True)

            if {
                "Sevk Tarihi",
                "Proforma Tarihi",
            }.issubset(teslim_edilenler.columns):
                teslim_edilenler["Gün Farkı"] = (
                    teslim_edilenler["Sevk Tarihi"]
                    - teslim_edilenler["Proforma Tarihi"]
                ).dt.days
            
            for kolon in [
                "Proforma Tarihi",
                "Termin Tarihi",
                "Sevk Tarihi",
                "Ulaşma Tarihi",
            ]:
                if kolon in teslim_edilenler.columns:
                    teslim_edilenler[kolon] = (
                        teslim_edilenler[kolon]
                        .dt.strftime("%d/%m/%Y")
                        .fillna("")
                        .replace({"NaT": ""})
                    )
            if "Gün Farkı" in teslim_edilenler.columns:
                teslim_edilenler["Gün Farkı"] = teslim_edilenler["Gün Farkı"].apply(
                    lambda x: "" if pd.isna(x) else int(x)
                )

            display_columns = [
                "Müşteri Adı",
                "Ülke",
                "Proforma No",
                "Proforma Tarihi",
                "Termin Tarihi",
                "Sevk Tarihi",
                "Ulaşma Tarihi",
                "Gün Farkı",
                "Tutar",
                "Açıklama",
            ]

            mevcut_kolonlar = [
                kolon for kolon in display_columns if kolon in teslim_edilenler.columns
            ]

            st.dataframe(
                teslim_edilenler[mevcut_kolonlar],
                use_container_width=True,
            )
        else:
            st.info("Teslim edilmiş sipariş yok.")
    else:
        st.info("Teslim edilmiş sipariş yok.")

    # ---- Vade Takibi Tablosu (HERKES GÖRÜR) ----
    st.markdown("### Vadeli Fatura ve Tahsilat Takibi")

    for col in ["Vade Tarihi", "Ödendi", "Ödenen Tutar"]:
        if col not in invoices_df.columns:
            if col == "Vade Tarihi":
                invoices_df[col] = ""
            elif col == "Ödendi":
                invoices_df[col] = False
            else:
                invoices_df[col] = 0.0

    invoices_df["Ödendi"] = invoices_df["Ödendi"].fillna(False).astype(bool)
    invoices_df["Ödenen Tutar"] = pd.to_numeric(
        invoices_df["Ödenen Tutar"], errors="coerce"
    ).fillna(0.0)
    kalan_serisi = (
        invoices_df["Tutar_num"].fillna(0.0) - invoices_df["Ödenen Tutar"]
    ).clip(lower=0.0)

    vade_ts = pd.to_datetime(invoices_df["Vade Tarihi"], errors="coerce")
    outstanding_mask = kalan_serisi > 0.01
    vadesi_gelmemis_m = (vade_ts > today_norm) & outstanding_mask
    vadesi_bugun_m = (vade_ts.dt.date == today_norm.date()) & outstanding_mask
    gecikmis_m = (vade_ts < today_norm) & outstanding_mask

    tg_sum = float(kalan_serisi.loc[vadesi_gelmemis_m].sum())
    tb_sum = float(kalan_serisi.loc[vadesi_bugun_m].sum())
    gec_sum = float(kalan_serisi.loc[gecikmis_m].sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Vadeleri Gelmeyen", f"{tg_sum:,.2f} USD", f"{int(vadesi_gelmemis_m.sum())} Fatura")
    c2.metric("Bugün Vadesi Dolan", f"{tb_sum:,.2f} USD", f"{int(vadesi_bugun_m.sum())} Fatura")
    c3.metric("Geciken Ödemeler", f"{gec_sum:,.2f} USD", f"{int(gecikmis_m.sum())} Fatura")

    for col in ["Proforma No", "Vade (gün)", "Ödendi", "Ülke", "Satış Temsilcisi", "Ödeme Şekli"]:
        if col not in invoices_df.columns:
            invoices_df[col] = "" if col != "Ödendi" else False
    invoices_df["Ödendi"] = invoices_df["Ödendi"].fillna(False).astype(bool)

    vade_df = invoices_df[
        invoices_df["Vade Tarihi"].notna() & outstanding_mask
    ].copy()
    gecikmis_df = pd.DataFrame()
    if vade_df.empty:
        st.info("Açık vade kaydı yok.")
    else:
        vade_df["Vade Tarihi"] = pd.to_datetime(vade_df["Vade Tarihi"])
        vade_df["Kalan Gün"] = (vade_df["Vade Tarihi"] - pd.to_datetime(datetime.date.today())).dt.days
        vade_df = vade_df.sort_values("Kalan Gün", ascending=True)
        gecikmis_df = vade_df[vade_df["Kalan Gün"] < 0].copy()
        vade_df_display = vade_df.copy()
        vade_df_display["Kalan Bakiye"] = vade_df_display.index.map(kalan_serisi.get)
        vade_df_display["Kalan Bakiye"] = vade_df_display["Kalan Bakiye"].fillna(0.0).map(lambda x: f"{x:,.2f} USD")
        st.dataframe(
            vade_df_display[["Müşteri Adı", "Ülke", "Fatura No", "Vade Tarihi", "Tutar", "Kalan Gün", "Kalan Bakiye"]],
            use_container_width=True,
        )

    st.markdown("#### Gecikmiş Ödemeler")
    if not gecikmis_df.empty:
        gecikmis_display = gecikmis_df.copy()
        gecikmis_display["Kalan Bakiye"] = gecikmis_display.index.map(kalan_serisi.get)
        gecikmis_display["Kalan Bakiye"] = gecikmis_display["Kalan Bakiye"].fillna(0.0).map(lambda x: f"{x:,.2f} USD")
        st.dataframe(
            gecikmis_display[["Müşteri Adı", "Ülke", "Fatura No", "Vade Tarihi", "Tutar", "Kalan Gün", "Kalan Bakiye"]],
            use_container_width=True,
        )
    else:
        st.info("Gecikmiş ödeme bulunmuyor.")
    st.markdown("### Satış Analitiği Özeti")

    summary_cols = st.columns(4)
    summary_cols[0].metric("Toplam Fatura Tutarı", f"{toplam_fatura_tutar:,.2f} USD")

    if date_col and invoices_df[date_col].notna().any():
        last_30_start = today_norm - pd.Timedelta(days=29)
        last_30_end = today_norm + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
        mask_last_30 = invoices_df[date_col].between(last_30_start, last_30_end, inclusive="both")
        last_30_total = float(invoices_df.loc[mask_last_30, "Tutar_num"].sum())
        summary_cols[1].metric("Son 30 Gün Cirosu", f"{last_30_total:,.2f} USD", f"{int(mask_last_30.sum())} Fatura")

        current_year = today_norm.year
        mask_year = invoices_df[date_col].dt.year == current_year
        year_total = float(invoices_df.loc[mask_year, "Tutar_num"].sum())
        summary_cols[2].metric(f"{current_year} Toplamı", f"{year_total:,.2f} USD", f"{int(mask_year.sum())} Fatura")
    else:
        summary_cols[1].metric("Son 30 Gün Cirosu", "0.00 USD")
        summary_cols[2].metric(f"{today_norm.year} Toplamı", "0.00 USD")

    if "Müşteri Adı" in invoices_df.columns:
        active_customers = (
            invoices_df["Müşteri Adı"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", np.nan)
            .dropna()
            .nunique()
        )
    else:
        active_customers = 0
    summary_cols[3].metric("Aktif Müşteri", str(active_customers))

    if invoices_df.empty:
        st.info("Satış analitiği için fatura kaydı bulunmuyor.")
    elif "Müşteri Adı" in invoices_df.columns:
        top_df = invoices_df[invoices_df["Tutar_num"] > 0].copy()
        top_df["Müşteri Adı"] = top_df["Müşteri Adı"].fillna("Bilinmeyen Müşteri").astype(str).str.strip()
        top_customers = (
            top_df.groupby("Müşteri Adı")["Tutar_num"]
            .sum()
            .sort_values(ascending=False)
            .head(5)
            .reset_index()
        )
        if not top_customers.empty:
            st.markdown("#### En Yüksek Ciroya Sahip İlk 5 Müşteri")
            display_df = top_customers.copy()
            display_df["Toplam Ciro"] = display_df["Tutar_num"].map(lambda x: f"{x:,.2f} USD")
            st.dataframe(display_df[["Müşteri Adı", "Toplam Ciro"]], use_container_width=True)
        else:
            st.info("Müşteri bazında ciro hesaplanacak veri bulunamadı.")
    else:
        st.info("Satış analitiği için müşteri bilgisi bulunmuyor.")


    st.markdown("<hr>", unsafe_allow_html=True)
    st.info("Daha detaylı işlem yapmak için sol menüden ilgili bölüme geçebilirsiniz.")


### ===========================
### === CARİ EKLEME MENÜSÜ ===
### ===========================

if menu == "Cari Kayıtlar":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Yeni Müşteri Ekle</h2>", unsafe_allow_html=True)

    # ---- Yardımcılar: doğrulama & normalizasyon ----
    import re
    def _clean_text(s):
        return (str(s or "")).strip()

    def _valid_email(s):
        s = _clean_text(s)
        if not s:
            return True  # boşsa zorunlu değil; doluysa kontrol
        # basit ve sağlam bir desen
        return re.match(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$", s) is not None

    def _normalized_phone(s):
        # sadece rakamları al, 10–15 haneye izin ver
        digits = re.sub(r"\D+", "", str(s or ""))
        return digits

    # Mükerrer kontrol için set (ad+ülke)
    if df_musteri.empty:
        existing_pairs = set()
    else:
        existing_pairs = set(
            (str(a).strip().lower(), str(u).strip().lower())
            for a, u in zip(df_musteri.get("Müşteri Adı", []), df_musteri.get("Ülke", []))
        )

    with st.form("add_customer", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("Müşteri Adı *", placeholder="Örn: ABC Dış Ticaret Ltd.")
            phone = st.text_input("Telefon", placeholder="+90 ...")
            email = st.text_input("E-posta", placeholder="ornek@firma.com")
            address = st.text_area("Adres")
            kategori = st.selectbox("Kategori", ["Avrupa bayi", "bayi", "müşteri", "yeni müşteri"], index=3)
            aktif_pasif = st.selectbox("Durum", ["Aktif", "Pasif"], index=0)
        with c2:
            ulke = st.selectbox("Ülke *", ulke_listesi)
            temsilci = st.selectbox("Satış Temsilcisi *", temsilci_listesi)
            vade_gun = st.number_input("Vade (Gün Sayısı)", min_value=0, max_value=365, value=0, step=1)
            odeme_sekli = st.selectbox("Ödeme Şekli", ["Peşin", "Mal Mukabili", "Vesaik Mukabili", "Akreditif", "Diğer"])
            para_birimi = st.selectbox("Para Birimi", ["USD", "EURO", "TL", "RUBLE"], index=0)
            dt_secim = st.selectbox("DT Seçin", ["DT-1", "DT-2", "DT-3", "DT-4"], index=0)

        col_submit1, col_submit2 = st.columns([1, 1])
        save_clicked = col_submit1.form_submit_button("Kaydet")
        save_and_send_clicked = col_submit2.form_submit_button("Kaydet ve Muhasebeye Gönder")


    if save_clicked or save_and_send_clicked:
        send_to_accounting = save_and_send_clicked
        # --- Normalizasyon ---
        name_n = _clean_text(name)
        ulke_n = _clean_text(ulke)
        email_n = _clean_text(email)
        phone_n = _normalized_phone(phone)

        # --- Zorunlu alanlar ---
        errors = []
        if not name_n:
            errors.append("Müşteri adı boş olamaz.")
        if not ulke_n:
            errors.append("Ülke seçimi zorunludur.")
        if not temsilci:
            errors.append("Satış temsilcisi seçimi zorunludur.")
        if not _valid_email(email_n):
            errors.append("E-posta formatı hatalı görünüyor.")

        # --- Mükerrer kontrol (Ad + Ülke) ---
        key = (name_n.lower(), ulke_n.lower())
        if key in existing_pairs:
            errors.append("Aynı ada ve ülkeye ait bir müşteri zaten kayıtlı görünüyor.")

        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        # --- Yeni satır ---
        new_row = {
            "Müşteri Adı": name_n,
            "Telefon": phone_n,                          # normalize edilmiş
            "E-posta": email_n,
            "Adres": _clean_text(address),
            "Ülke": ulke_n,
            "Satış Temsilcisi": temsilci,
            "Kategori": kategori,
            "Durum": aktif_pasif,
            "Vade (Gün)": vade_gun,
            "Ödeme Şekli": odeme_sekli,
            "Para Birimi": para_birimi,
            "DT Seçimi": dt_secim,
            "Oluşturma Tarihi": datetime.date.today(),  # faydalı meta
        }

        # --- Kaydet ---
        df_musteri = pd.concat([df_musteri, pd.DataFrame([new_row])], ignore_index=True)
        update_excel()

        if send_to_accounting:
            # --- Muhasebeye e-posta (sende tanımlı yardımcılar) ---
            try:
                yeni_cari_txt_olustur(new_row)
                attachments = []
                try:
                    with open("yeni_cari.txt", "rb") as attachment_file:
                        attachments.append(("yeni_cari.txt", attachment_file.read(), "text/plain"))
                except FileNotFoundError:
                    pass

                send_email(
                    to_email=["muhasebe@sekeroglugroup.com", "h.boy@sekeroglugroup.com"],
                    subject="Yeni Cari Açılışı",
                    body="Muhasebe için yeni cari açılışı ekte gönderilmiştir.",
                    attachments=attachments,
                    fallback_txt_path="yeni_cari.txt",
                )
                st.success("Müşteri eklendi ve e‑posta ile muhasebeye gönderildi!")
            except Exception as e:
                st.warning(f"Müşteri eklendi ancak e‑posta gönderilemedi: {e}")
        else:
            st.success("Müşteri eklendi. Muhasebeye göndermek için ilgili butonu kullanabilirsiniz.")

        st.balloons()
        st.rerun()



                
### ===========================
### === MÜŞTERİ LİSTESİ MENÜSÜ (Cloud-Sağlam) ===
### ===========================

# — Zorunlu sütunları garanti altına al —
gerekli_kolonlar = [
    "ID", "Müşteri Adı", "Telefon", "E-posta", "Adres",
    "Ülke", "Satış Temsilcisi", "Kategori", "Durum",
    "Vade (Gün)", "Ödeme Şekli", "Para Birimi", "DT Seçimi"
]
for col in gerekli_kolonlar:
    if col not in df_musteri.columns:
        if col == "ID":
            # eksikse tüm satırlar için üret
            if len(df_musteri) > 0:
                df_musteri[col] = [str(uuid.uuid4()) for _ in range(len(df_musteri))]
            else:
                df_musteri[col] = []
        elif col == "Vade (Gün)":
            df_musteri[col] = ""
        else:
            df_musteri[col] = ""

# — Eski kayıtlarda ID boşsa doldur —
mask_id_bos = df_musteri["ID"].isna() | (df_musteri["ID"].astype(str).str.strip() == "")
if mask_id_bos.any():
    df_musteri.loc[mask_id_bos, "ID"] = [str(uuid.uuid4()) for _ in range(mask_id_bos.sum())]
    update_excel()

if menu == "Müşteri Portföyü":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Müşteri Listesi</h2>", unsafe_allow_html=True)

    # ---- Üst Araçlar: Arama + Filtreler ----
    with st.container():
        c1, c2, c3, c4 = st.columns([2, 1.2, 1.2, 1.2])
        aranacak = c1.text_input("Arama (Ad / Telefon / E-posta / Adres)", value="")
        ulke_filtre = c2.multiselect("Ülke Filtresi", sorted([u for u in df_musteri["Ülke"].dropna().unique() if str(u).strip()]), default=[])
        temsilci_filtre = c3.multiselect("Temsilci Filtresi", sorted([t for t in df_musteri["Satış Temsilcisi"].dropna().unique() if str(t).strip()]), default=[])
        durum_filtre = c4.multiselect("Durum", ["Aktif", "Pasif"], default=["Aktif"])  # Varsayılan: Aktif

    # ---- Filtreleme mantığı ----
    view_df = df_musteri.copy()

    # Durum filtresi
    if len(durum_filtre) > 0:
        view_df = view_df[view_df["Durum"].isin(durum_filtre)]

    # Ülke filtresi
    if len(ulke_filtre) > 0:
        view_df = view_df[view_df["Ülke"].isin(ulke_filtre)]

    # Temsilci filtresi
    if len(temsilci_filtre) > 0:
        view_df = view_df[view_df["Satış Temsilcisi"].isin(temsilci_filtre)]

    # Arama filtresi
    if aranacak.strip():
        s = aranacak.strip().lower()
        def _match(row):
            fields = [
                row.get("Müşteri Adı", ""), row.get("Telefon", ""), row.get("E-posta", ""),
                row.get("Adres", ""), row.get("Ülke", ""), row.get("Satış Temsilcisi", "")
            ]
            return any(s in str(x).lower() for x in fields)
        view_df = view_df[view_df.apply(_match, axis=1)]

    # Görüntü tablosu (boşları sadece tabloda “—” yap)
    show_cols = ["Müşteri Adı", "Ülke", "Satış Temsilcisi", "Telefon", "E-posta", "Adres", "Kategori", "Durum", "Vade (Gün)", "Ödeme Şekli", "Para Birimi", "DT Seçimi"]
    for c in show_cols:
        if c not in view_df.columns:
            view_df[c] = ""

    table_df = view_df[show_cols].replace({np.nan: "—", "": "—"})
    table_df = table_df.sort_values("Müşteri Adı").reset_index(drop=True)

    # Özet bilgi ve dışa aktar
    top_row = st.columns([3, 1])
    with top_row[0]:
        st.markdown(f"<div style='color:#219A41; font-weight:700;'>Toplam Kayıt: {len(view_df)}</div>", unsafe_allow_html=True)
    with top_row[1]:
        st.download_button(
            "CSV indir",
            data=table_df.to_csv(index=False).encode("utf-8"),
            file_name="musteri_listesi.csv",
            mime="text/csv",
            use_container_width=True
        )

    if table_df.empty:
        st.markdown("<div style='color:#b00020; font-weight:bold; font-size:1.1em;'>Kayıt bulunamadı.</div>", unsafe_allow_html=True)
    else:
        st.dataframe(table_df, use_container_width=True)

        st.markdown("#### Toplu Mail Gönderimi")
    with st.expander("Filtrelenmiş müşterilere toplu mail gönder", expanded=False):
        email_pattern = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
        mevcut_epostalar = []
        if "E-posta" in view_df.columns:
            mevcut_epostalar = [
                str(e).strip()
                for e in view_df["E-posta"].dropna().tolist()
                if str(e).strip()
            ]

        mevcut_epostalar = sorted({e for e in mevcut_epostalar if email_pattern.match(e)})

        if not mevcut_epostalar:
            st.info("Seçili filtrelere göre geçerli e-posta adresi bulunamadı.")
        else:
            with st.form("bulk_mail_form"):
                secilen_adresler = st.multiselect(
                    "Alıcıları seçin", mevcut_epostalar, help="Listelenen geçerli adreslerden seçim yapın."
                )
                ek_adresler = st.text_area(
                    "Ek alıcılar", "", help="Virgül, satır sonu veya noktalı virgülle ayırabilirsiniz."
                )
                toplu_konu = st.text_input("Konu")
                toplu_icerik = st.text_area("İçerik")
                yuklenen_dosyalar = st.file_uploader(
                    "Ekler", accept_multiple_files=True, help="E-posta eklerini buradan yükleyin."
                )
                gonder = st.form_submit_button("E-postayı Gönder")

            if gonder:
                tum_adresler = list(secilen_adresler)
                if ek_adresler.strip():
                    for parca in re.split(r"[\s,;]+", ek_adresler):
                        aday = parca.strip()
                        if aday and email_pattern.match(aday):
                            tum_adresler.append(aday)

                # Sıra koruyarak tekrarları kaldır
                benzersiz_adresler = []
                for adres in tum_adresler:
                    if adres not in benzersiz_adresler:
                        benzersiz_adresler.append(adres)

                if not benzersiz_adresler:
                    st.warning("Lütfen en az bir geçerli alıcı belirtin.")
                elif not toplu_konu.strip():
                    st.warning("Konu alanı boş bırakılamaz.")
                else:
                    ekler = []
                    if yuklenen_dosyalar:
                        for dosya in yuklenen_dosyalar:
                            try:
                                icerik = dosya.read()
                            except Exception:
                                icerik = dosya.getvalue()
                            ekler.append(
                                (
                                    dosya.name,
                                    icerik,
                                    dosya.type or "application/octet-stream",
                                )
                            )

                    try:
                        send_email(
                            to_email=benzersiz_adresler,
                            subject=toplu_konu.strip(),
                            body=toplu_icerik,
                            attachments=ekler,
                        )
                    except Exception as e:
                        st.error(f"E-posta gönderimi sırasında hata oluştu: {e}")
                    else:
                        st.success("E-postalar başarıyla gönderildi.")


    st.markdown("<h4 style='margin-top: 24px;'>Müşteri Düzenle / Sil</h4>", unsafe_allow_html=True)

    # Düzenleme/Silme için seçim: ID ile — güvenli
    # Önce ekranda gösterilen view_df'ten seçim yaptırıyoruz (alfabetik)
    secenek_df = view_df.sort_values("Müşteri Adı").reset_index(drop=True)
    if secenek_df.empty:
        st.info("Düzenlemek/silmek için uygun kayıt yok.")
    else:
        secim = st.selectbox(
            "Düzenlenecek Müşteriyi Seçin",
            options=secenek_df["ID"].tolist(),
            format_func=lambda _id: f"{secenek_df.loc[secenek_df['ID']==_id, 'Müşteri Adı'].values[0]} ({secenek_df.loc[secenek_df['ID']==_id, 'Kategori'].values[0]})"
        )

        # Orijinal index (ana df_musteri içinden) — ID ile eşle
        orj_mask = (df_musteri["ID"] == secim)
        if not orj_mask.any():
            st.warning("Beklenmeyen hata: Seçilen kayıt ana tabloda bulunamadı.")
        else:
            orj_idx = df_musteri.index[orj_mask][0]

            with st.form("edit_existing_customer"):
                name = st.text_input("Müşteri Adı", value=str(df_musteri.at[orj_idx, "Müşteri Adı"]))
                phone = st.text_input("Telefon", value=str(df_musteri.at[orj_idx, "Telefon"]))
                email = st.text_input("E-posta", value=str(df_musteri.at[orj_idx, "E-posta"]))
                address = st.text_area("Adres", value=str(df_musteri.at[orj_idx, "Adres"]))

                # Ülke / Temsilci seçimleri
                try:
                    ulke_def = df_musteri.at[orj_idx, "Ülke"]
                    ulke_idx = ulke_listesi.index(ulke_def) if ulke_def in ulke_listesi else 0
                except Exception:
                    ulke_idx = 0
                ulke = st.selectbox("Ülke", ulke_listesi, index=ulke_idx)

                try:
                    tem_def = df_musteri.at[orj_idx, "Satış Temsilcisi"]
                    tem_idx = temsilci_listesi.index(tem_def) if tem_def in temsilci_listesi else 0
                except Exception:
                    tem_idx = 0
                temsilci = st.selectbox("Satış Temsilcisi", temsilci_listesi, index=tem_idx)

                kategori = st.selectbox(
                    "Kategori",
                    sorted(["Avrupa bayi", "bayi", "müşteri", "yeni müşteri"]),
                    index=sorted(["Avrupa bayi", "bayi", "müşteri", "yeni müşteri"]).index(df_musteri.at[orj_idx, "Kategori"])
                    if df_musteri.at[orj_idx, "Kategori"] in ["Avrupa bayi", "bayi", "müşteri", "yeni müşteri"] else 0
                )
                aktif_pasif = st.selectbox(
                    "Durum", ["Aktif", "Pasif"],
                    index=(0 if str(df_musteri.at[orj_idx, "Durum"]) == "Aktif" else 1)
                )

                vade = st.text_input("Vade (Gün)", value=str(df_musteri.at[orj_idx, "Vade (Gün)"]) if "Vade (Gün)" in df_musteri.columns else "")
                odeme_sekli = st.selectbox(
                    "Ödeme Şekli",
                    ["Peşin", "Mal Mukabili", "Vesaik Mukabili", "Akreditif", "Diğer"],
                    index=["Peşin", "Mal Mukabili", "Vesaik Mukabili", "Akreditif", "Diğer"].index(df_musteri.at[orj_idx, "Ödeme Şekli"])
                    if df_musteri.at[orj_idx, "Ödeme Şekli"] in ["Peşin", "Mal Mukabili", "Vesaik Mukabili", "Akreditif", "Diğer"] else 0
                )

                para_birimi = st.selectbox(
                    "Para Birimi",
                    ["EURO", "USD", "TL", "RUBLE"],
                    index=["EURO", "USD", "TL", "RUBLE"].index(df_musteri.at[orj_idx, "Para Birimi"]) if df_musteri.at[orj_idx, "Para Birimi"] in ["EURO", "USD", "TL", "RUBLE"] else 0
                )

                dt_secimi = st.selectbox(
                    "DT Seçimi",
                    ["DT-1", "DT-2", "DT-3", "DT-4"],
                    index=["DT-1", "DT-2", "DT-3", "DT-4"].index(df_musteri.at[orj_idx, "DT Seçimi"]) if df_musteri.at[orj_idx, "DT Seçimi"] in ["DT-1", "DT-2", "DT-3", "DT-4"] else 0
                )

                colu, colm, cols = st.columns(3)
                guncelle = colu.form_submit_button("Güncelle")
                muhasebe_gonder = colm.form_submit_button("Muhasebeye Gönder")
                sil = cols.form_submit_button("Sil")

            if guncelle:
                df_musteri.at[orj_idx, "Müşteri Adı"] = name
                df_musteri.at[orj_idx, "Telefon"] = phone
                df_musteri.at[orj_idx, "E-posta"] = email
                df_musteri.at[orj_idx, "Adres"] = address
                df_musteri.at[orj_idx, "Ülke"] = ulke
                df_musteri.at[orj_idx, "Satış Temsilcisi"] = temsilci
                df_musteri.at[orj_idx, "Kategori"] = kategori
                df_musteri.at[orj_idx, "Durum"] = aktif_pasif
                df_musteri.at[orj_idx, "Vade (Gün)"] = vade
                df_musteri.at[orj_idx, "Ödeme Şekli"] = odeme_sekli
                df_musteri.at[orj_idx, "Para Birimi"] = para_birimi
                df_musteri.at[orj_idx, "DT Seçimi"] = dt_secimi
                update_excel()
                st.success("Müşteri bilgisi güncellendi!")
                st.rerun()

            if muhasebe_gonder:
                guncel_bilgiler = {
                    "Müşteri Adı": name,
                    "Telefon": phone,
                    "E-posta": email,
                    "Adres": address,
                    "Ülke": ulke,
                    "Satış Temsilcisi": temsilci,
                    "Kategori": kategori,
                    "Durum": aktif_pasif,
                    "Vade (Gün)": vade,
                    "Ödeme Şekli": odeme_sekli,
                    "Para Birimi": para_birimi,
                    "DT Seçimi": dt_secimi,
                }

                try:
                    yeni_cari_txt_olustur(guncel_bilgiler)
                    attachments = []
                    try:
                        with open("yeni_cari.txt", "rb") as attachment_file:
                            attachments.append(("yeni_cari.txt", attachment_file.read(), "text/plain"))
                    except FileNotFoundError:
                        pass

                    send_email(
                        to_email=["muhasebe@sekeroglugroup.com", "h.boy@sekeroglugroup.com"],
                        subject="Güncel Cari Bilgisi",
                        body="Mevcut müşteri için güncel cari bilgileri ekte yer almaktadır.",
                        attachments=attachments,
                        fallback_txt_path="yeni_cari.txt",
                    )
                except Exception as e:
                    st.warning(f"Muhasebeye gönderim sırasında bir hata oluştu: {e}")
                else:
                    st.success("Güncel müşteri bilgileri muhasebeye gönderildi!")
                    st.rerun()

            if sil:
                df_musteri = df_musteri.drop(orj_idx).reset_index(drop=True)
                update_excel()
                st.success("Müşteri kaydı silindi!")
                st.rerun()


### ===========================
### === ETKİLEŞİM GÜNLÜĞÜ (Cloud-Sağlam) ===
### ===========================

# Zorunlu kolonlar
gerekli = ["ID", "Müşteri Adı", "Tarih", "Tip", "Açıklama"]
for c in gerekli:
    if c not in df_kayit.columns:
        df_kayit[c] = ""

# Eski kayıtlarda ID yoksa doldur
mask_bos_id = df_kayit["ID"].isna() | (df_kayit["ID"].astype(str).str.strip() == "")
if mask_bos_id.any():
    df_kayit.loc[mask_bos_id, "ID"] = [str(uuid.uuid4()) for _ in range(mask_bos_id.sum())]
    update_excel()

if menu == "Etkileşim Günlüğü":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Etkileşim Günlüğü</h2>", unsafe_allow_html=True)
    
    st.subheader("Kayıt Ekranı")
    secim = st.radio("Lütfen işlem seçin:", ["Yeni Kayıt", "Eski Kayıt", "Tarih Aralığı ile Kayıtlar"], horizontal=False)

    # --- Ortak: müşteri listesi (boş hariç, alfabetik) ---
    musteri_options = [""] + sorted([
        m for m in df_musteri["Müşteri Adı"].dropna().unique()
        if isinstance(m, str) and m.strip() != ""
    ])

    # === YENİ KAYIT ===
    if secim == "Yeni Kayıt":
        with st.form("add_kayit"):
            musteri_sec = st.selectbox("Müşteri Seç", musteri_options, index=0)
            tarih = st.date_input("Tarih", value=datetime.date.today(), format="DD/MM/YYYY")
            tip = st.selectbox("Tip", ["Arama", "Görüşme", "Ziyaret"])
            aciklama = st.text_area("Açıklama")
            submitted = st.form_submit_button("Kaydet")
            if submitted:
                if not musteri_sec:
                    st.error("Lütfen bir müşteri seçiniz.")
                else:
                    new_row = {
                        "ID": str(uuid.uuid4()),
                        "Müşteri Adı": musteri_sec,
                        "Tarih": tarih,
                        "Tip": tip,
                        "Açıklama": aciklama
                    }
                    df_kayit = pd.concat([df_kayit, pd.DataFrame([new_row])], ignore_index=True)
                    update_excel()
                    st.success("Kayıt eklendi!")
                    st.rerun()

    # === ESKİ KAYIT (Listele / Ara / Düzenle / Sil) ===
    elif secim == "Eski Kayıt":
        colf1, colf2, colf3 = st.columns([2, 1, 1])
        musteri_f = colf1.selectbox("Müşteri Filtresi", ["(Hepsi)"] + sorted(df_kayit["Müşteri Adı"].dropna().unique().tolist()))
        tip_f = colf2.multiselect("Tip Filtresi", ["Arama", "Görüşme", "Ziyaret"], default=[])
        aranacak = colf3.text_input("Ara (açıklama)", value="")

        view = df_kayit.copy()
        # Filtreler
        if musteri_f and musteri_f != "(Hepsi)":
            view = view[view["Müşteri Adı"] == musteri_f]
        if tip_f:
            view = view[view["Tip"].isin(tip_f)]
        if aranacak.strip():
            s = aranacak.lower().strip()
            view = view[view["Açıklama"].astype(str).str.lower().str.contains(s, na=False)]

        # Görünüm tablosu
        if not view.empty:
            goster = view.copy()
            goster["Tarih"] = pd.to_datetime(goster["Tarih"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(goster[["Müşteri Adı", "Tarih", "Tip", "Açıklama"]].sort_values("Tarih", ascending=False), use_container_width=True)

            # Dışa aktar
            st.download_button(
                "CSV indir",
                data=goster.to_csv(index=False).encode("utf-8"),
                file_name="gorusme_kayitlari.csv",
                mime="text/csv"
            )
        else:
            st.info("Seçilen filtrelere uygun kayıt bulunamadı.")

        # Düzenleme / Silme
        st.markdown("#### Kayıt Düzenle / Sil")
        if view.empty:
            st.caption("Önce filtreleriyle bir kayıt listeleyin.")
        else:
            # Seçim ID ile (en son ekleneni üste almak için tarihe göre sıralayalım)
            view_sorted = view.sort_values("Tarih", ascending=False).reset_index(drop=True)
            sec_id = st.selectbox(
                "Kayıt Seçin",
                options=view_sorted["ID"].tolist(),
                format_func=lambda _id: f"{view_sorted.loc[view_sorted['ID']==_id, 'Müşteri Adı'].values[0]} | {view_sorted.loc[view_sorted['ID']==_id, 'Tip'].values[0]}"
            )

            # Orijinal index
            orj_mask = (df_kayit["ID"] == sec_id)
            if not orj_mask.any():
                st.warning("Beklenmeyen hata: Kayıt ana tabloda bulunamadı.")
            else:
                orj_idx = df_kayit.index[orj_mask][0]
                with st.form("edit_kayit"):
                    musteri_g = st.selectbox("Müşteri", musteri_options, index=(musteri_options.index(df_kayit.at[orj_idx, "Müşteri Adı"]) if df_kayit.at[orj_idx, "Müşteri Adı"] in musteri_options else 0))
                    try:
                        tarih_g = pd.to_datetime(df_kayit.at[orj_idx, "Tarih"]).date()
                    except Exception:
                        tarih_g = datetime.date.today()
                    tarih_g = st.date_input("Tarih", value=tarih_g, format="DD/MM/YYYY")
                    tip_g = st.selectbox("Tip", ["Arama", "Görüşme", "Ziyaret"], index=["Arama","Görüşme","Ziyaret"].index(df_kayit.at[orj_idx,"Tip"]) if df_kayit.at[orj_idx,"Tip"] in ["Arama","Görüşme","Ziyaret"] else 0)
                    aciklama_g = st.text_area("Açıklama", value=str(df_kayit.at[orj_idx, "Açıklama"]))
                    colu, cols = st.columns(2)
                    guncelle = colu.form_submit_button("Güncelle")
                    sil = cols.form_submit_button("Sil")

                if guncelle:
                    df_kayit.at[orj_idx, "Müşteri Adı"] = musteri_g
                    df_kayit.at[orj_idx, "Tarih"] = tarih_g
                    df_kayit.at[orj_idx, "Tip"] = tip_g
                    df_kayit.at[orj_idx, "Açıklama"] = aciklama_g
                    update_excel()
                    st.success("Kayıt güncellendi!")
                    st.rerun()

                if sil:
                    df_kayit = df_kayit.drop(orj_idx).reset_index(drop=True)
                    update_excel()
                    st.success("Kayıt silindi!")
                    st.rerun()

    # === TARİH ARALIĞI İLE KAYITLAR ===
    elif secim == "Tarih Aralığı ile Kayıtlar":
        col1, col2 = st.columns(2)
        with col1:
            baslangic = st.date_input("Başlangıç Tarihi", value=datetime.date.today() - datetime.timedelta(days=7), format="DD/MM/YYYY")
        with col2:
            bitis = st.date_input("Bitiş Tarihi", value=datetime.date.today(), format="DD/MM/YYYY")

        # Sağlam tarih filtrelemesi
        tser = pd.to_datetime(df_kayit["Tarih"], errors="coerce")
        start_ts = pd.to_datetime(baslangic)
        end_ts = pd.to_datetime(bitis) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
        tarih_arasi = df_kayit[tser.between(start_ts, end_ts, inclusive="both")].copy()

        if not tarih_arasi.empty:
            goster = tarih_arasi.copy()
            goster["Tarih"] = pd.to_datetime(goster["Tarih"], errors="coerce").dt.strftime('%d/%m/%Y')
            st.dataframe(goster.sort_values("Tarih", ascending=False), use_container_width=True)
            st.download_button(
                "CSV indir",
                data=goster.to_csv(index=False).encode("utf-8"),
                file_name="gorusme_kayitlari_tarih_araligi.csv",
                mime="text/csv"
            )
        else:
            st.info("Bu tarihler arasında kayıt yok.")

elif menu == "Özel Gün Tebrikleri":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Özel Gün Tebrikleri</h2>", unsafe_allow_html=True)
    st.markdown("Seçilmiş kişilere bayram ve yeni yıl tebrik e-postaları gönderebilirsiniz.")
    
    st.markdown(
        "<h4 style='margin-top:0;'>Seçilmiş kişilere bayram ve yeni yıl tebrikleri gönderin</h4>",
        unsafe_allow_html=True,
    )

    musteriden_eposta = extract_unique_emails(df_musteri.get("E-posta")) if "E-posta" in df_musteri.columns else []
    fuardan_eposta = extract_unique_emails(df_fuar_musteri.get("E-mail")) if "E-mail" in df_fuar_musteri.columns else []

    unknown_country_label = "(Belirtilmedi)"

    email_country_map = {}

    def kayitlardan_ulke_ekle(df, email_col, country_col):
        if df is None or df.empty or email_col not in df.columns:
            return
        country_available = country_col in df.columns
        for _, row in df.iterrows():
            raw_email = row.get(email_col, "")
            if pd.isna(raw_email) or str(raw_email).strip() == "":
                continue
            parsed_emails = extract_unique_emails(pd.Series([raw_email]))
            country_val = str(row.get(country_col, "") if country_available else "").strip()
            if not country_val or country_val.lower() in ["nan", "none"]:
                country_val = unknown_country_label
            for mail in parsed_emails:
                email_country_map.setdefault(mail, set()).add(country_val)

    kayitlardan_ulke_ekle(df_musteri, "E-posta", "Ülke")
    kayitlardan_ulke_ekle(df_fuar_musteri, "E-mail", "Ülke")

    tum_epostalar = sorted(email_country_map.keys(), key=lambda x: x.lower())
    tum_ulkeler = sorted({ulke for ulkeler in email_country_map.values() for ulke in ulkeler})

    if not tum_epostalar:
        st.info("Gönderim yapabileceğiniz e-posta adresi bulunamadı.")
        tum_ulkeler = []
    else:
        st.caption(f"Toplam benzersiz e-posta: {len(tum_epostalar)}")

    st.session_state.setdefault("bulk_mail_subject", "")
    st.session_state.setdefault("bulk_mail_body", "")
    st.session_state.setdefault("bulk_mail_template_info", None)
    st.session_state.setdefault("holiday_selected_template", "(Şablon seçiniz)")
    st.session_state.setdefault("holiday_selected_language", None)

    secili_ulkeler = st.multiselect(
        "Ülke filtresi",
        tum_ulkeler,
        default=tum_ulkeler,
        help="Sadece seçilen ülkelerle ilişkili e-posta adresleri listelenir.",
    ) if tum_ulkeler else []

    secili_ulkeler_kumesi = set(secili_ulkeler) if secili_ulkeler else set(tum_ulkeler)

    filtrelenmis_epostalar = [
        mail
        for mail, ulkeler in email_country_map.items()
        if not secili_ulkeler_kumesi or ulkeler.intersection(secili_ulkeler_kumesi)
    ]
    filtrelenmis_epostalar = sorted(filtrelenmis_epostalar, key=lambda x: x.lower())

    tumunu_sec_opsiyonu = "(Tümünü seç)"
    multiselect_options = ([tumunu_sec_opsiyonu] + filtrelenmis_epostalar) if filtrelenmis_epostalar else []
    varsayilan_secim = [tumunu_sec_opsiyonu] if filtrelenmis_epostalar else []

    secilen_epostalar_raw = st.multiselect(
        "E-posta adresleri",
        multiselect_options,
        default=varsayilan_secim,
        help="Gönderim yapılacak adresleri seçin.",
    ) if multiselect_options else []

    if tumunu_sec_opsiyonu in secilen_epostalar_raw:
        secilen_epostalar = filtrelenmis_epostalar
    else:
        secilen_epostalar = secilen_epostalar_raw

    derived_languages = sorted({
        COUNTRY_LANGUAGE_MAP.get(country)
        for country in secili_ulkeler_kumesi
        if country != unknown_country_label and COUNTRY_LANGUAGE_MAP.get(country)
    })

    template_placeholder = "(Şablon seçiniz)"
    template_options = [template_placeholder] + list(HOLIDAY_MAIL_TEMPLATES.keys())
    if st.session_state["holiday_selected_template"] not in template_options:
        st.session_state["holiday_selected_template"] = template_placeholder

    secilen_sablon = st.selectbox(
        "Şablon",
        template_options,
        key="holiday_selected_template",
    )        


    secilen_dil = None
    fallback_kullanildi = False
    if secilen_sablon != template_placeholder:
        sablon_dilleri = sorted(HOLIDAY_MAIL_TEMPLATES.get(secilen_sablon, {}).keys())
        aday_diller = [lang for lang in sablon_dilleri if lang in derived_languages]
        if not aday_diller:
            aday_diller = sablon_dilleri

        if len(aday_diller) == 1:
            secilen_dil = aday_diller[0]
            st.session_state["holiday_selected_language"] = secilen_dil
            st.caption(
                f"Seçilen ülkelere göre dil otomatik olarak {LANGUAGE_LABELS.get(secilen_dil, secilen_dil)} olarak belirlendi."
            )
        else:
            onceki_dil = st.session_state.get("holiday_selected_language")
            varsayilan_dil = None
            if derived_languages:
                for dil in derived_languages:
                    if dil in aday_diller:
                        varsayilan_dil = dil
                        break
            if varsayilan_dil is None:
                varsayilan_dil = aday_diller[0] if aday_diller else None
            if onceki_dil not in aday_diller:
                st.session_state["holiday_selected_language"] = varsayilan_dil

            secilen_dil = st.selectbox(
                "Dil",
                aday_diller,
                format_func=lambda lang: LANGUAGE_LABELS.get(lang, lang.upper()),
                key="holiday_selected_language",                      
            
            )
            
        mevcut_info = st.session_state.get("bulk_mail_template_info")
        onceki_konu = st.session_state.get("bulk_mail_subject", "")
        onceki_govde = st.session_state.get("bulk_mail_body", "")
        mevcut_sablon_adi = mevcut_info["name"] if mevcut_info else None
        mevcut_sablon_dili = mevcut_info["language"] if mevcut_info else None
        alanlar_bos = not onceki_konu.strip() and not onceki_govde.strip()
        alanlar_mevcut_sablona_esit = bool(
            mevcut_info
            and onceki_konu == mevcut_info.get("subject", "")
            and onceki_govde == mevcut_info.get("body", "")
        )
            
        yeni_icerik, fallback_kullanildi = get_holiday_template_content(secilen_sablon, secilen_dil)
        if yeni_icerik:
            sablon_degisti = (
                mevcut_sablon_adi != secilen_sablon
                or mevcut_sablon_dili != secilen_dil
            )
            manuel_duzenleme = bool(mevcut_info and not alanlar_mevcut_sablona_esit)

            sablon_yuklenecek = False
            if alanlar_bos:
                sablon_yuklenecek = True
            elif sablon_degisti and not manuel_duzenleme:
                sablon_yuklenecek = True
            elif not sablon_degisti and alanlar_mevcut_sablona_esit:
                sablon_yuklenecek = True

            if sablon_yuklenecek:
                st.session_state["bulk_mail_subject"] = yeni_icerik["subject"]
                st.session_state["bulk_mail_body"] = yeni_icerik["body"]
                st.session_state["bulk_mail_template_info"] = {
                    "name": secilen_sablon,
                    "language": secilen_dil,
                    "subject": yeni_icerik["subject"],
                    "body": yeni_icerik["body"],
                    "fallback": fallback_kullanildi,
                }
    else:
        st.session_state["holiday_selected_language"] = None


    konu = st.text_input("E-posta Konusu", key="bulk_mail_subject")
    govde = st.text_area(
        "HTML Gövde",
        key="bulk_mail_body",
        height=280,
        help="İsterseniz metni Türkçe/İngilizce olarak düzenleyebilirsiniz.",
    )

    yuklenen_gorsel = st.file_uploader(
        "Görsel ekleyin (isteğe bağlı)",
        type=["png", "jpg", "jpeg", "gif", "webp", "svg"],
        accept_multiple_files=False,
        help="Tek bir görsel yükleyebilirsiniz. Görsel inline gönderim için saklanacaktır.",
    )

    if filtrelenmis_epostalar:
        onizleme_df = pd.DataFrame(
            [
                {
                    "E-posta": mail,
                    "Ülkeler": ", ".join(sorted(email_country_map.get(mail, {unknown_country_label}))),
                }
                for mail in filtrelenmis_epostalar
            ]
         )
        
        st.dataframe(onizleme_df, use_container_width=True, hide_index=True)

    etkin_ulke_text = "Tüm ülkeler" if not secili_ulkeler else ", ".join(secili_ulkeler)
    st.markdown(
        f"<div style='margin-top:12px; font-size:0.95em;'>"
        f"<strong>Aktif ülke filtresi:</strong> {etiket if (etiket := etkin_ulke_text) else 'Tüm ülkeler'}<br>"
        f"<strong>Seçilen adres sayısı:</strong> {len(secilen_epostalar)}<br>"
        "<strong>Not:</strong> Gönderimlerde varsayılan HTML imzası otomatik olarak eklenecektir."
        "</div>",
        unsafe_allow_html=True,
    )

    if st.button("Toplu Maili Gönder", type="primary"):
        if not secilen_epostalar:
            st.warning("Lütfen en az bir e-posta adresi seçiniz.")
        elif not konu.strip():
            st.warning("Lütfen e-posta konusu giriniz.")
        else:
            attachments = [yuklenen_gorsel] if yuklenen_gorsel else []
            try:
                with st.spinner("E-postalar gönderiliyor..."):
                    send_fair_bulk_email(secilen_epostalar, konu.strip(), govde, attachments=attachments)
                st.success(f"E-posta {len(secilen_epostalar)} alıcıya başarıyla gönderildi.")
            except Exception as exc:
                st.error(f"Gönderim sırasında bir hata oluştu: {exc}")
       


### ===========================
### --- TEKLİF YÖNETİMİ (Cloud-Sağlam) ---
### ===========================

elif menu == "Teklif Yönetimi":

    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Teklif Yönetimi</h2>", unsafe_allow_html=True)

    # --- Zorunlu kolonlar + ID backfill ---
    gerekli = ["ID", "Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Ürün/Hizmet", "Açıklama", "Durum", "PDF"]
    for c in gerekli:
        if c not in df_teklif.columns:
            df_teklif[c] = ""
    mask_bos_id = df_teklif["ID"].astype(str).str.strip().isin(["", "nan"])
    if mask_bos_id.any():
        df_teklif.loc[mask_bos_id, "ID"] = [str(uuid.uuid4()) for _ in range(mask_bos_id.sum())]
        update_excel()

    # --- Otomatik teklif no ---
    def otomatik_teklif_no():
        if df_teklif.empty or "Teklif No" not in df_teklif.columns:
            return "TKF-0001"
        sayilar = pd.to_numeric(
            df_teklif["Teklif No"].astype(str).str.extract(r'(\d+)$')[0], errors='coerce'
        ).dropna().astype(int)
        yeni_no = (sayilar.max() + 1) if not sayilar.empty else 1
        return f"TKF-{yeni_no:04d}"


    # ---------- ÜST ÖZET: Açık teklifler ----------
    tkg = df_teklif.copy()
    tkg["Tarih"] = pd.to_datetime(tkg["Tarih"], errors="coerce")
    acik_teklifler = tkg[tkg["Durum"] == "Açık"].sort_values(["Müşteri Adı", "Teklif No"])
    toplam_teklif = float(acik_teklifler["Tutar"].apply(smart_to_num).sum())
    acik_teklif_sayi = len(acik_teklifler)
    st.subheader("Açık Pozisyondaki Teklifler")
    st.markdown(
        f"<div style='font-size:1.05em; color:#11998e; font-weight:bold;'>Toplam: {toplam_teklif:,.2f} USD | "
        f"Toplam Açık Teklif: {acik_teklif_sayi} adet</div>",
        unsafe_allow_html=True
    )
    if not acik_teklifler.empty:
        goster = acik_teklifler.copy()
        goster["Tarih"] = goster["Tarih"].dt.strftime("%d/%m/%Y")
        st.dataframe(goster[["Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Ürün/Hizmet", "Açıklama"]], use_container_width=True)
    else:
        st.info("Açık teklif bulunmuyor.")

    st.markdown("##### Lütfen bir işlem seçin")
    col1, col2 = st.columns(2)
    with col1:
        yeni_teklif_buton = st.button("Yeni Teklif")
    with col2:
        eski_teklif_buton = st.button("Eski Teklifler / Düzenle")

    if "teklif_view" not in st.session_state:
        st.session_state['teklif_view'] = None
    if yeni_teklif_buton:
        st.session_state['teklif_view'] = "yeni"
    if eski_teklif_buton:
        st.session_state['teklif_view'] = "eski"

    # ============== YENİ TEKLİF ==============
    if st.session_state['teklif_view'] == "yeni":
        musteri_list = [""] + sorted(df_musteri["Müşteri Adı"].dropna().unique().tolist())
        st.subheader("Yeni Teklif Ekle")
        with st.form("add_teklif"):
            musteri_sec = st.selectbox("Müşteri Seç", musteri_list, key="yeni_teklif_musteri")
            tarih = st.date_input("Tarih", value=datetime.date.today(), format="DD/MM/YYYY")
            teklif_no = st.text_input("Teklif No", value=otomatik_teklif_no())
            tutar = st.text_input("Tutar (USD)")
            urun = st.text_input("Ürün/Hizmet")
            aciklama = st.text_area("Açıklama")
            durum = st.selectbox("Durum", ["Açık", "Sonuçlandı", "Beklemede"])
            pdf_file = st.file_uploader("Teklif PDF", type="pdf")
            submitted = st.form_submit_button("Kaydet")

            pdf_link = ""
            if submitted:
                if not teklif_no.strip():
                    st.error("Teklif No boş olamaz!")
                elif not musteri_sec:
                    st.error("Lütfen müşteri seçiniz!")
                else:
                    # PDF'yi Drive'a yükle (varsa)
                    if pdf_file:
                        temiz_musteri = "".join(x if x.isalnum() else "_" for x in str(musteri_sec))
                        temiz_tarih = str(tarih).replace("-", "")
                        pdf_filename = f"{temiz_musteri}__{temiz_tarih}__{teklif_no}.pdf"
                        temp_path = os.path.join(".", pdf_filename)
                        with open(temp_path, "wb") as f:
                            f.write(pdf_file.read())
                        gfile = drive.CreateFile({'title': pdf_filename, 'parents': [{'id': FIYAT_TEKLIFI_ID}]})
                        gfile.SetContentFile(temp_path)
                        gfile.Upload()
                        pdf_link = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
                        güvenli_sil(temp_path)

                    new_row = {
                        "ID": str(uuid.uuid4()),
                        "Müşteri Adı": musteri_sec,
                        "Tarih": tarih,
                        "Teklif No": teklif_no,
                        "Tutar": tutar,
                        "Ürün/Hizmet": urun,
                        "Açıklama": aciklama,
                        "Durum": durum,
                        "PDF": pdf_link
                    }
                    df_teklif = pd.concat([df_teklif, pd.DataFrame([new_row])], ignore_index=True)
                    update_excel()
                    st.success("Teklif eklendi!")
                    st.session_state['teklif_view'] = None
                    st.rerun()

    # ============== ESKİ TEKLİFLER / DÜZENLE / SİL ==============
    if st.session_state['teklif_view'] == "eski":
        st.subheader("Eski Teklifler")

        # ---- Filtreler ----
        f1, f2, f3, f4 = st.columns([1.5, 1, 1.3, 1.2])
        musteri_f = f1.selectbox("Müşteri", ["(Hepsi)"] + sorted(df_teklif["Müşteri Adı"].dropna().unique().tolist()))
        durum_f = f2.multiselect("Durum", ["Açık", "Beklemede", "Sonuçlandı"], default=[])
        # Tarih aralığı
        tmp = pd.to_datetime(df_teklif["Tarih"], errors="coerce")
        min_dt = (tmp.min().date() if tmp.notna().any() else datetime.date.today())
        max_dt = (tmp.max().date() if tmp.notna().any() else datetime.date.today())
        d1 = f3.date_input("Başlangıç", value=min_dt)
        d2 = f4.date_input("Bitiş", value=max_dt)
        aranacak = st.text_input("Ara (ürün/açıklama/teklif no)")

        view = df_teklif.copy()
        view["Tarih"] = pd.to_datetime(view["Tarih"], errors="coerce")

        if musteri_f and musteri_f != "(Hepsi)":
            view = view[view["Müşteri Adı"] == musteri_f]
        if durum_f:
            view = view[view["Durum"].isin(durum_f)]

        start_ts = pd.to_datetime(d1)
        end_ts = pd.to_datetime(d2) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
        view = view[view["Tarih"].between(start_ts, end_ts, inclusive="both")]

        if aranacak.strip():
            s = aranacak.lower().strip()
            view = view[
                view["Ürün/Hizmet"].astype(str).str.lower().str.contains(s, na=False) |
                view["Açıklama"].astype(str).str.lower().str.contains(s, na=False) |
                view["Teklif No"].astype(str).str.lower().str.contains(s, na=False)
            ]

        # Toplam ve tablo
        toplam_view = float(view["Tutar"].apply(smart_to_num).sum())
        st.markdown(f"<div style='margin:.25rem 0 .5rem 0; font-weight:600;'>Filtreli Toplam: {toplam_view:,.2f} USD</div>", unsafe_allow_html=True)

        if not view.empty:
            tablo = view.sort_values("Tarih", ascending=False).copy()
            tablo["Tarih"] = tablo["Tarih"].dt.strftime("%d/%m/%Y")
            st.dataframe(tablo[["Müşteri Adı", "Tarih", "Teklif No", "Tutar", "Durum", "Ürün/Hizmet", "Açıklama"]], use_container_width=True)
            st.download_button(
                "CSV indir",
                data=tablo.to_csv(index=False).encode("utf-8"),
                file_name="teklifler.csv",
                mime="text/csv"
            )
        else:
            st.info("Filtrelere göre teklif bulunamadı.")

        # ---- Düzenle / Sil ----
        st.markdown("#### Teklif Düzenle / Sil")
        if view.empty:
            st.caption("Önce filtrelerle bir kayıt listeleyin.")
        else:
            v_sorted = view.sort_values("Tarih", ascending=False).reset_index(drop=True)
            sec_id = st.selectbox(
                "Teklif Seçiniz",
                options=v_sorted["ID"].tolist(),
                format_func=lambda _id: f"{v_sorted.loc[v_sorted['ID']==_id, 'Müşteri Adı'].values[0]} | {v_sorted.loc[v_sorted['ID']==_id, 'Teklif No'].values[0]}"
            )

            orj_mask = (df_teklif["ID"] == sec_id)
            if not orj_mask.any():
                st.warning("Beklenmeyen hata: Teklif ana tabloda bulunamadı.")
            else:
                orj_idx = df_teklif.index[orj_mask][0]
                # Var olan PDF linkini göster
                mevcut_pdf = str(df_teklif.at[orj_idx, "PDF"]) if pd.notna(df_teklif.at[orj_idx, "PDF"]) else ""
                if mevcut_pdf:
                    st.markdown(f"**Mevcut PDF:** [Görüntüle]({mevcut_pdf})", unsafe_allow_html=True)

                with st.form("edit_teklif"):
                    try:
                        tarih_g = pd.to_datetime(df_teklif.at[orj_idx, "Tarih"]).date()
                    except Exception:
                        tarih_g = datetime.date.today()
                    tarih_g = st.date_input("Tarih", value=tarih_g, format="DD/MM/YYYY")
                    teklif_no_g = st.text_input("Teklif No", value=str(df_teklif.at[orj_idx, "Teklif No"]))
                    musteri_g = st.selectbox(
                        "Müşteri",
                        [""] + sorted(df_musteri["Müşteri Adı"].dropna().unique().tolist()),
                        index=([""] + sorted(df_musteri["Müşteri Adı"].dropna().unique().tolist())).index(df_teklif.at[orj_idx, "Müşteri Adı"]) if df_teklif.at[orj_idx, "Müşteri Adı"] in ([""] + sorted(df_musteri["Müşteri Adı"].dropna().unique().tolist())) else 0
                    )
                    tutar_g = st.text_input("Tutar (USD)", value=str(df_teklif.at[orj_idx, "Tutar"]))
                    urun_g = st.text_input("Ürün/Hizmet", value=str(df_teklif.at[orj_idx, "Ürün/Hizmet"]))
                    aciklama_g = st.text_area("Açıklama", value=str(df_teklif.at[orj_idx, "Açıklama"]))
                    durum_g = st.selectbox("Durum", ["Açık", "Beklemede", "Sonuçlandı"],
                                           index=["Açık", "Beklemede", "Sonuçlandı"].index(df_teklif.at[orj_idx, "Durum"]) if df_teklif.at[orj_idx, "Durum"] in ["Açık","Beklemede","Sonuçlandı"] else 0)
                    pdf_yeni = st.file_uploader("PDF Güncelle (opsiyonel)", type="pdf", key=f"pdf_guncel_{sec_id}")
                    colu, cols = st.columns(2)
                    guncelle = colu.form_submit_button("Güncelle")
                    sil = cols.form_submit_button("Sil")

                if guncelle:
                    pdf_link_final = mevcut_pdf
                    if pdf_yeni:
                        # Yeni PDF yükle
                        temiz_m = "".join(x if x.isalnum() else "_" for x in str(musteri_g or "musteri"))
                        temiz_t = str(tarih_g).replace("-", "")
                        fname = f"{temiz_m}__{temiz_t}__{teklif_no_g}.pdf"
                        tmp_path = os.path.join(".", fname)
                        with open(tmp_path, "wb") as f:
                            f.write(pdf_yeni.read())
                        gfile = drive.CreateFile({'title': fname, 'parents': [{'id': FIYAT_TEKLIFI_ID}]})
                        gfile.SetContentFile(tmp_path)
                        gfile.Upload()
                        pdf_link_final = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
                        güvenli_sil(tmp_path)

                    df_teklif.at[orj_idx, "Tarih"] = tarih_g
                    df_teklif.at[orj_idx, "Teklif No"] = teklif_no_g
                    df_teklif.at[orj_idx, "Müşteri Adı"] = musteri_g
                    df_teklif.at[orj_idx, "Tutar"] = tutar_g
                    df_teklif.at[orj_idx, "Ürün/Hizmet"] = urun_g
                    df_teklif.at[orj_idx, "Açıklama"] = aciklama_g
                    df_teklif.at[orj_idx, "Durum"] = durum_g
                    df_teklif.at[orj_idx, "PDF"] = pdf_link_final
                    update_excel()
                    st.success("Teklif güncellendi!")
                    st.rerun()

                if sil:
                    df_teklif = df_teklif.drop(orj_idx).reset_index(drop=True)
                    update_excel()
                    st.success("Teklif silindi!")
                    st.rerun()




### ===========================
### --- PROFORMA TAKİBİ MENÜSÜ (Cloud-Sağlam) ---
### ===========================

elif menu == "Proforma Yönetimi":

    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Proforma Yönetimi</h2>", unsafe_allow_html=True)

    if "convert_proforma_id" not in st.session_state:
        st.session_state.convert_proforma_id = None
          
    # ---- Drive klasör ID'leri (üstten tanımlıysa onları, yoksa EVRAK_KLASOR_ID'yi kullan) ----
    PROFORMA_PDF_FOLDER_ID   = globals().get("PROFORMA_PDF_FOLDER_ID", globals().get("EVRAK_KLASOR_ID"))
    SIPARIS_FORMU_FOLDER_ID  = globals().get("SIPARIS_FORMU_FOLDER_ID", globals().get("EVRAK_KLASOR_ID"))

    # ---- Kolon güvenliği + ID backfill ----
    gerekli = ["ID","Müşteri Adı","Tarih","Proforma No","Tutar","Açıklama","Durum","PDF",
               "Vade (gün)","Sevk Durumu","Ülke","Satış Temsilcisi","Ödeme Şekli","Termin Tarihi","Ulaşma Tarihi","Sipariş Formu"]
    for c in gerekli:
        if c not in df_proforma.columns:
            df_proforma[c] = ""
    mask_bos_id = df_proforma["ID"].astype(str).str.strip().isin(["","nan"])
    if mask_bos_id.any():
        df_proforma.loc[mask_bos_id, "ID"] = [str(uuid.uuid4()) for _ in range(mask_bos_id.sum())]
        update_excel()

    def render_siparis_formu_yukleme(df, hedef_id):
        if not hedef_id:
            return

        hedef_mask = df["ID"] == hedef_id
        if not hedef_mask.any():
            st.session_state.convert_proforma_id = None
            return

        hedef_idx = df.index[hedef_mask][0]
        hedef_kayit = df.loc[hedef_idx]

        st.markdown("#### Siparişe Dönüştürme - Sipariş Formu Yükle")
        st.info(
            f"{hedef_kayit['Müşteri Adı']} - {hedef_kayit['Proforma No']} için sipariş formunu yükleyin."
        )

        form_key = f"siparis_formu_upload_{hedef_id}"
        with st.form(form_key):
            siparis_formu_file = st.file_uploader(
                "Sipariş Formu PDF", type="pdf", key=f"sf_{hedef_id}"
            )
            col_sf1, col_sf2 = st.columns(2)
            kaydet_sf = col_sf1.form_submit_button("Sipariş Formunu Kaydet ve Dönüştür")
            vazgec_sf = col_sf2.form_submit_button("Vazgeç")

        if kaydet_sf:
            if siparis_formu_file is None:
                st.error("Sipariş formu yüklenmeli.")
                return

            sf_name = (
                f"{hedef_kayit['Müşteri Adı']}_{hedef_kayit['Proforma No']}_SiparisFormu_"
                f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            )
            tmp = os.path.join(".", sf_name)
            with open(tmp, "wb") as f:
                f.write(siparis_formu_file.read())

            gfile = drive.CreateFile({'title': sf_name, 'parents': [{'id': SIPARIS_FORMU_FOLDER_ID}]})
            gfile.SetContentFile(tmp)
            gfile.Upload()
            sf_url = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
            güvenli_sil(tmp)

            df.at[hedef_idx, "Sipariş Formu"] = sf_url
            df.at[hedef_idx, "Durum"] = "Siparişe Dönüştü"
            df.at[hedef_idx, "Sevk Durumu"] = ""
            st.session_state.convert_proforma_id = None
            update_excel()
            st.success(
                "Sipariş formu kaydedildi ve durum 'Siparişe Dönüştü' olarak güncellendi!"
            )
            st.rerun()

        elif vazgec_sf:
            st.session_state.convert_proforma_id = None
            st.info("Siparişe dönüştürme işlemi iptal edildi.")
            st.rerun()

        update_excel()
   
    
    # ---------- ÜST ÖZET: Bekleyen Proformalar ----------
    pview = df_proforma.copy()
    pview["Tarih"] = pd.to_datetime(pview["Tarih"], errors="coerce")
    beklemede_kayitlar = pview[pview["Durum"] == "Beklemede"].sort_values(["Tarih","Müşteri Adı"], ascending=[False, True])
    toplam_bekleyen = float(beklemede_kayitlar["Tutar"].apply(smart_to_num).sum())

    st.subheader("Bekleyen Proformalar")
    st.markdown(f"<div style='font-weight:600;'>Toplam Bekleyen: {toplam_bekleyen:,.2f} USD</div>", unsafe_allow_html=True)
    if not beklemede_kayitlar.empty:
        g = beklemede_kayitlar.copy()
        g["Tarih"] = g["Tarih"].dt.strftime("%d/%m/%Y")
        st.dataframe(g[["Müşteri Adı","Proforma No","Tarih","Tutar","Durum","Vade (gün)","Sevk Durumu"]], use_container_width=True)
    else:
        st.info("Beklemede proforma bulunmuyor.")

    # ---------- Müşteri seçimi ----------
    musteri_list = sorted([x for x in df_musteri["Müşteri Adı"].dropna().unique() if str(x).strip()!=""]) if not df_musteri.empty else []
    musteri_sec = st.selectbox("Müşteri Seç", [""] + musteri_list)

    if musteri_sec:
        st.write("Proforma işlemi seçin:")
        islem = st.radio("", ["Yeni Kayıt","Eski Kayıt / Düzenle"], horizontal=True)

        # ============== YENİ KAYIT ==============
        if islem == "Yeni Kayıt":
            musteri_info = df_musteri[df_musteri["Müşteri Adı"] == musteri_sec]
            default_ulke      = musteri_info["Ülke"].values[0] if not musteri_info.empty else ""
            default_temsilci  = musteri_info["Satış Temsilcisi"].values[0] if not musteri_info.empty else ""
            default_odeme     = musteri_info["Ödeme Şekli"].values[0] if not musteri_info.empty else ""

            with st.form("add_proforma"):
                tarih      = st.date_input("Tarih", value=datetime.date.today())
                proforma_no= st.text_input("Proforma No")
                tutar      = st.text_input("Tutar (USD)")
                vade_gun   = st.text_input("Vade (gün)")
                ulke       = st.text_input("Ülke", value=default_ulke, disabled=True)
                temsilci   = st.text_input("Satış Temsilcisi", value=default_temsilci, disabled=True)
                odeme      = st.text_input("Ödeme Şekli", value=default_odeme, disabled=True)
                aciklama   = st.text_area("Açıklama")
                durum      = st.selectbox("Durum", ["Beklemede","İptal","Faturası Kesildi","Siparişe Dönüştü"], index=0)
                pdf_file   = st.file_uploader("Proforma PDF", type="pdf")
                submitted  = st.form_submit_button("Kaydet")

                if submitted:
                    if not proforma_no.strip() or not vade_gun.strip():
                        st.error("Proforma No ve Vade (gün) boş olamaz!")
                    else:
                        # Aynı müşteri+proforma no duplike kontrolü
                        if ((df_proforma["Müşteri Adı"]==musteri_sec) & (df_proforma["Proforma No"].astype(str)==proforma_no.strip())).any():
                            st.warning("Bu Proforma No bu müşteri için zaten kayıtlı.")
                        else:
                            pdf_link = ""
                            if pdf_file and PROFORMA_PDF_FOLDER_ID:
                                fname = f"{musteri_sec}_{tarih}_{proforma_no}.pdf"
                                tmp = os.path.join(".", fname)
                                with open(tmp,"wb") as f: f.write(pdf_file.read())
                                gfile = drive.CreateFile({'title': fname, 'parents':[{'id': PROFORMA_PDF_FOLDER_ID}]})
                                gfile.SetContentFile(tmp); gfile.Upload()
                                pdf_link = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
                                güvenli_sil(tmp)

                            new_row = {
                                "ID": str(uuid.uuid4()),
                                "Müşteri Adı": musteri_sec,
                                "Tarih": tarih,
                                "Proforma No": proforma_no.strip(),
                                "Tutar": tutar,
                                "Vade (gün)": vade_gun,
                                "Ülke": default_ulke,
                                "Satış Temsilcisi": default_temsilci,
                                "Ödeme Şekli": default_odeme,
                                "Açıklama": aciklama,
                                "Durum": "Beklemede" if durum!="Siparişe Dönüştü" else "Siparişe Dönüştü",
                                "PDF": pdf_link,
                                "Sipariş Formu": "",
                                "Sevk Durumu": "",
                                "Termin Tarihi": "",
                                "Ulaşma Tarihi": ""
                            }
                            df_proforma = pd.concat([df_proforma, pd.DataFrame([new_row])], ignore_index=True)
                            update_excel()
                            st.success("Proforma eklendi!")
                            st.rerun()

        # ============== ESKİ KAYIT / DÜZENLE / SİL / SİPARİŞE DÖNÜŞTÜR ==============
        elif islem == "Eski Kayıt / Düzenle":
            # Seçilen müşterinin beklemede + diğer durumları
            kayitlar = df_proforma[df_proforma["Müşteri Adı"] == musteri_sec].copy()
            if (
                st.session_state.convert_proforma_id
                and st.session_state.convert_proforma_id not in kayitlar["ID"].tolist()
            ):
                st.session_state.convert_proforma_id = None
            if kayitlar.empty:
                st.info("Bu müşteriye ait proforma kaydı yok.")
            else:
                kayitlar["Tarih"] = pd.to_datetime(kayitlar["Tarih"], errors="coerce")
                st.dataframe(
                    kayitlar.sort_values("Tarih", ascending=False)[
                        ["Müşteri Adı","Proforma No","Tarih","Tutar","Durum","Vade (gün)","Sevk Durumu"]
                    ],
                    use_container_width=True
                )

                sec_id = st.selectbox(
                    "Proforma Seç",
                    options=kayitlar["ID"].tolist(),
                    format_func=lambda _id: f"{kayitlar.loc[kayitlar['ID']==_id,'Proforma No'].values[0]} | {kayitlar.loc[kayitlar['ID']==_id,'Tarih'].dt.strftime('%d/%m/%Y').values[0] if pd.notna(kayitlar.loc[kayitlar['ID']==_id,'Tarih'].values[0]) else ''}"
                )

                orj_mask = (df_proforma["ID"] == sec_id)
                if not orj_mask.any():
                    st.warning("Beklenmeyen hata: Kayıt bulunamadı.")
                else:
                    idx = df_proforma.index[orj_mask][0]
                    kayit = df_proforma.loc[idx]

                    if str(kayit.get("PDF","")).strip():
                        st.markdown(f"**Proforma PDF:** [Görüntüle]({kayit['PDF']})", unsafe_allow_html=True)

                    with st.form("edit_proforma"):
                        tarih_      = st.date_input("Tarih", value=(pd.to_datetime(kayit["Tarih"], errors="coerce").date() if pd.notna(pd.to_datetime(kayit["Tarih"], errors="coerce")) else datetime.date.today()))
                        proforma_no_= st.text_input("Proforma No", value=str(kayit["Proforma No"]))
                        tutar_      = st.text_input("Tutar (USD)", value=str(kayit["Tutar"]))
                        vade_gun_   = st.text_input("Vade (gün)", value=str(kayit["Vade (gün)"]))
                        aciklama_   = st.text_area("Açıklama", value=str(kayit["Açıklama"]))
                        durum_      = st.selectbox("Durum", ["Beklemede","Siparişe Dönüştü","İptal","Faturası Kesildi"],
                                                    index=["Beklemede","Siparişe Dönüştü","İptal","Faturası Kesildi"].index(kayit["Durum"]) if kayit["Durum"] in ["Beklemede","Siparişe Dönüştü","İptal","Faturası Kesildi"] else 0)
                        termin_     = st.date_input("Termin Tarihi", value=(pd.to_datetime(kayit.get("Termin Tarihi",""), errors="coerce").date() if pd.notna(pd.to_datetime(kayit.get("Termin Tarihi",""), errors="coerce")) else datetime.date.today()), key="termin_inp")
                        pdf_yeni    = st.file_uploader("Proforma PDF (güncelle - opsiyonel)", type="pdf")
                        colu, colm, cols = st.columns(3)
                        guncelle = colu.form_submit_button("Güncelle")
                        donustur = colm.form_submit_button("Siparişe Dönüştür (+ Sipariş Formu)")
                        sil      = cols.form_submit_button("Sil")

                    # --- GÜNCELLE ---
                    if guncelle:
                        pdf_final = str(kayit.get("PDF",""))
                        if pdf_yeni and PROFORMA_PDF_FOLDER_ID:
                            fname = f"{musteri_sec}_{tarih_}_{proforma_no_}.pdf"
                            tmp = os.path.join(".", fname)
                            with open(tmp,"wb") as f: f.write(pdf_yeni.read())
                            gfile = drive.CreateFile({'title': fname, 'parents':[{'id': PROFORMA_PDF_FOLDER_ID}]})
                            gfile.SetContentFile(tmp); gfile.Upload()
                            pdf_final = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
                            güvenli_sil(tmp)

                        df_proforma.at[idx, "Tarih"] = tarih_
                        df_proforma.at[idx, "Proforma No"] = proforma_no_
                        df_proforma.at[idx, "Tutar"] = tutar_
                        df_proforma.at[idx, "Vade (gün)"] = vade_gun_
                        df_proforma.at[idx, "Açıklama"] = aciklama_
                        df_proforma.at[idx, "Durum"] = durum_ if durum_ != "Siparişe Dönüştü" else df_proforma.at[idx, "Durum"]
                        df_proforma.at[idx, "Termin Tarihi"] = termin_
                        df_proforma.at[idx, "PDF"] = pdf_final
                        st.session_state.convert_proforma_id = None
                        update_excel()
                        st.success("Proforma güncellendi!")
                        st.rerun()

                    # --- SİPARİŞE DÖNÜŞTÜR (Sipariş Formu zorunlu) ---
                    if donustur:
                        st.session_state.convert_proforma_id = sec_id
                        st.rerun()
                        
                    # --- SİL ---
                    if sil:
                        st.session_state.convert_proforma_id = None
                        df_proforma = df_proforma.drop(idx).reset_index(drop=True)
                        update_excel()
                        st.success("Kayıt silindi!")
                        st.rerun()

                    render_siparis_formu_yukleme(df_proforma, st.session_state.convert_proforma_id)

                   
### ===========================
### --- SİPARİŞ OPERASYONLARI (ID tabanlı) ---
### ===========================

elif menu == "Sipariş Operasyonları":
    import uuid

    st.header("Güncel Sipariş Durumu")

    # ---- Kolon güvenliği + ID backfill ----
    gerekli = ["ID","Sevk Durumu","Termin Tarihi","Sipariş Formu","Ülke","Satış Temsilcisi","Ödeme Şekli","PDF","Durum","Tarih","Tutar","Müşteri Adı","Proforma No","Açıklama"]
    for c in gerekli:
        if c not in df_proforma.columns:
            df_proforma[c] = ""
    bos_id = df_proforma["ID"].astype(str).str.strip().isin(["","nan"])
    if bos_id.any():
        df_proforma.loc[bos_id, "ID"] = [str(uuid.uuid4()) for _ in range(bos_id.sum())]
        update_excel()

    # ---- Filtre: Siparişe dönmüş ama sevk edilmemiş/ulaşmamış kayıtlar
    siparisler = df_proforma[
        (df_proforma["Durum"] == "Siparişe Dönüştü")
        & (~df_proforma["Sevk Durumu"].isin(["Sevkedildi","Ulaşıldı"]))
    ].copy()

    if siparisler.empty:
        st.info("Henüz sevk edilmeyi bekleyen sipariş yok.")
        st.stop()

    # ---- Sıralama: Termin Tarihi
    siparisler["Termin Tarihi Order"] = pd.to_datetime(siparisler["Termin Tarihi"], errors="coerce")
    siparisler["Tarih"] = pd.to_datetime(siparisler["Tarih"], errors="coerce")
    siparisler = siparisler.sort_values(["Termin Tarihi Order","Tarih"], ascending=[True, True])

    # ---- Görünüm için format
    siparis_sayisi = len(siparisler)
    g = siparisler.copy()
    g["Tarih"] = g["Tarih"].dt.strftime("%d/%m/%Y")
    g["Termin Tarihi"] = pd.to_datetime(g["Termin Tarihi"], errors="coerce").dt.strftime("%d/%m/%Y")

    st.markdown(
        f"<h4 style='color:#219A41; font-weight:bold;'>Tüm Siparişe Dönüşenler ({siparis_sayisi} Adet)</h4>",
        unsafe_allow_html=True,
    )
    goruntulenecek_kolonlar = [
        "Tarih",
        "Müşteri Adı",
        "Termin Tarihi",
        "Ülke",
        "Satış Temsilcisi",
        "Ödeme Şekli",
        "Proforma No",
        "Tutar",
        "Açıklama",
    ]
    g_tab = g[goruntulenecek_kolonlar].reset_index(drop=True)
    g_tab.index = g_tab.index + 1
    g_tab.index.name = "Sıra"
    st.dataframe(g_tab, use_container_width=True)

    # ================= Termin Tarihi Güncelle =================
    st.markdown("#### Termin Tarihi Güncelle")
    sec_id_termin = st.selectbox(
        "Termin Tarihi Girilecek Sipariş",
        options=siparisler["ID"].tolist(),
        format_func=lambda _id: f"{siparisler.loc[siparisler['ID']==_id, 'Müşteri Adı'].values[0]} - {siparisler.loc[siparisler['ID']==_id, 'Proforma No'].values[0]}"
    )
    mask_termin = (df_proforma["ID"] == sec_id_termin)
    mevcut_termin = pd.to_datetime(df_proforma.loc[mask_termin, "Termin Tarihi"].values[0] if mask_termin.any() else None, errors="coerce")
    default_termin = (mevcut_termin.date() if pd.notna(mevcut_termin) else datetime.date.today())
    yeni_termin = st.date_input("Termin Tarihi", value=default_termin, key="termin_input")

    if st.button("Termin Tarihini Kaydet"):
        df_proforma.loc[mask_termin, "Termin Tarihi"] = yeni_termin
        update_excel()
        st.success("Termin tarihi kaydedildi!")
        st.rerun()

    # ================= Sevk Et (ETA’ya gönder) =================
    st.markdown("#### Siparişi Sevk Et (ETA İzleme Kaydına Gönder)")
    sec_id_sevk = st.selectbox(
        "Sevk Edilecek Sipariş",
        options=siparisler["ID"].tolist(),
        format_func=lambda _id: f"{siparisler.loc[siparisler['ID']==_id, 'Müşteri Adı'].values[0]} - {siparisler.loc[siparisler['ID']==_id, 'Proforma No'].values[0]}",
        key="sevk_sec"
    )
    if st.button("Sevkedildi → ETA İzlemeye Ekle"):
        # Proforma'dan bilgiler
        row = df_proforma.loc[df_proforma["ID"] == sec_id_sevk].iloc[0]
        # ETA kolon güvenliği
        for col in ETA_COLUMNS:        
            if col not in df_eta.columns:
                df_eta[col] = ""
        # ETA'ya ekle (varsa güncelleme)
        filt = (df_eta["Müşteri Adı"] == row["Müşteri Adı"]) & (df_eta["Proforma No"] == row["Proforma No"])
        if filt.any():
            df_eta.loc[filt, "Sevk Tarihi"] = row.get("Sevk Tarihi", "")           
            df_eta.loc[filt, "Açıklama"] = row.get("Açıklama","")
        else:
            sevk_tarih = row.get("Sevk Tarihi", "")            
            df_eta = pd.concat([df_eta, pd.DataFrame([{
                "Müşteri Adı": row["Müşteri Adı"],
                "Proforma No": row["Proforma No"],
                "Sevk Tarihi": sevk_tarih,               
                "ETA Tarihi": "",
                "Açıklama": row.get("Açıklama","")
            }])], ignore_index=True)
        # Proforma'yı işaretle
        df_proforma.loc[df_proforma["ID"] == sec_id_sevk, "Sevk Durumu"] = "Sevkedildi"
        update_excel()
        st.success("Sipariş sevkedildi ve ETA takibine gönderildi!")
        st.rerun()

    # ================= Beklemeye Al (Geri Çağır) =================
    st.markdown("#### Siparişi Beklemeye Al (Geri Çağır)")
    sec_id_geri = st.selectbox(
        "Beklemeye Alınacak Sipariş",
        options=siparisler["ID"].tolist(),
        format_func=lambda _id: f"{siparisler.loc[siparisler['ID']==_id, 'Müşteri Adı'].values[0]} - {siparisler.loc[siparisler['ID']==_id, 'Proforma No'].values[0]}",
        key="geri_sec"
    )
    if st.button("Beklemeye Al / Geri Çağır"):
        m = (df_proforma["ID"] == sec_id_geri)
        df_proforma.loc[m, ["Durum","Sevk Durumu","Termin Tarihi"]] = ["Beklemede","",""]
        update_excel()
        st.success("Sipariş tekrar bekleyen proformalar listesine alındı!")
        st.rerun()

    # ================= Linkler + Toplam =================
    st.markdown("#### Tıklanabilir Proforma ve Sipariş Formu Linkleri")
    for _, r in siparisler.iterrows():
        links = []
        if str(r.get("PDF","")).strip():
            links.append(f"[Proforma PDF: {r['Proforma No']}]({r['PDF']})")
        if str(r.get("Sipariş Formu","")).strip():
            fname = f"{r['Müşteri Adı']}__{r['Proforma No']}__SiparisFormu"
            links.append(f"[Sipariş Formu: {fname}]({r['Sipariş Formu']})")
        if links:
            st.markdown(" - " + " | ".join(links), unsafe_allow_html=True)

    # Toplam bekleyen sevk tutarı (akıllı parse)

    toplam = float(siparisler["Tutar"].apply(smart_to_num).sum())
    st.markdown(f"<div style='color:#219A41; font-weight:bold;'>*Toplam Bekleyen Sevk: {toplam:,.2f} USD*</div>", unsafe_allow_html=True)

### ===========================
### --- İHRACAT EVRAKLARI MENÜSÜ ---
### ===========================

elif menu == "Fatura işlemleri":

    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Fatura işlemleri</h2>", unsafe_allow_html=True)

    # ---- Sütun güvenliği + benzersiz ID ----
    gerekli_kolonlar = [
        "ID","Müşteri Adı","Proforma No","Fatura No","Fatura Tarihi","Tutar","Ödenen Tutar",
        "Vade (gün)","Vade Tarihi","Ülke","Satış Temsilcisi","Ödeme Şekli",
        "Commercial Invoice","Sağlık Sertifikası","Packing List","Konşimento",
        "İhracat Beyannamesi","Fatura PDF","Sipariş Formu","Yük Resimleri",
        "EK Belgeler","Ödendi"
    ]
    for c in gerekli_kolonlar:
        if c not in df_evrak.columns:
            if c == "Ödendi":
                df_evrak[c] = False
            elif c == "Ödenen Tutar":
                df_evrak[c] = 0.0
            else:
                df_evrak[c] = ""

    bos_id_mask = df_evrak["ID"].astype(str).str.strip().isin(["","nan"])
    if bos_id_mask.any():
        df_evrak.loc[bos_id_mask, "ID"] = [str(uuid.uuid4()) for _ in range(bos_id_mask.sum())]
        update_excel()

        # ---- Otomatik seçim için session state anahtarları ----
    musteri_key = "invoice_customer_select"
    proforma_key = "invoice_proforma_select"
    pending_select_key = "pending_invoice_select"
    pending_reset_flag_key = "pending_invoice_select_reset"
    st.session_state.setdefault(musteri_key, "")
    st.session_state.setdefault(proforma_key, "")
    st.session_state.setdefault("invoice_last_customer", "")
    st.session_state.setdefault(pending_select_key, "")
    st.session_state.setdefault(pending_reset_flag_key, False)

    if "Sevk Durumu" not in df_proforma.columns:
        df_proforma["Sevk Durumu"] = ""
    if "Durum" not in df_proforma.columns:
        df_proforma["Durum"] = ""

    # ---- Sevk edilmiş fakat faturası kesilmemiş siparişler ----
    st.markdown("### Faturası Kesilmemiş Sevkli Siparişler")
    sevkedilen_mask = df_proforma["Sevk Durumu"] == "Sevkedildi"
    if "Durum" in df_proforma.columns:
        sevkedilen_mask &= df_proforma["Durum"].astype(str).str.strip().eq("Siparişe Dönüştü")
    pending_orders = df_proforma[sevkedilen_mask].copy()

    if not pending_orders.empty:
        pending_orders["Termin Tarihi Sıra"] = pd.to_datetime(
            pending_orders.get("Termin Tarihi"), errors="coerce"
        )
        pending_orders["Tarih Sıra"] = pd.to_datetime(
            pending_orders.get("Tarih"), errors="coerce"
        )
        pending_orders = pending_orders.sort_values(
            ["Termin Tarihi Sıra", "Tarih Sıra"], ascending=[True, True]
        )

    if not pending_orders.empty:
        pending_orders["ID"] = pending_orders["ID"].astype(str)
        pending_orders["Müşteri Adı"] = pending_orders["Müşteri Adı"].astype(str)
        pending_orders["Proforma No"] = pending_orders["Proforma No"].astype(str)

        pending_orders = pending_orders[
            pending_orders["ID"].str.strip() != ""
        ]

        if not df_evrak.empty:
            invoice_pairs = set(
                (
                    str(m).strip().lower(),
                    str(p).strip().lower(),
                )
                for m, p in zip(df_evrak.get("Müşteri Adı", []), df_evrak.get("Proforma No", []))
                if str(m).strip() or str(p).strip()
            )
            pending_orders = pending_orders[
                ~pending_orders.apply(
                    lambda r: (
                        str(r.get("Müşteri Adı", "")).strip().lower(),
                        str(r.get("Proforma No", "")).strip().lower(),
                    )
                    in invoice_pairs,
                    axis=1,
                )
            ]

        pending_orders = pending_orders[
            pending_orders["Proforma No"].astype(str).str.strip() != ""
        ]

    if pending_orders.empty:
        st.info("Sevk edilip henüz faturası kaydedilmemiş sipariş bulunmuyor.")
    else:
        display_cols = [
            "ID",
            "Müşteri Adı",
            "Proforma No",
            "Termin Tarihi",
            "Tutar",
            "Açıklama",
        ]
        for col in display_cols:
            if col not in pending_orders.columns:
                pending_orders[col] = ""
        table = pending_orders[display_cols].copy()
        table["Termin Tarihi"] = pd.to_datetime(table["Termin Tarihi"], errors="coerce").dt.strftime("%d/%m/%Y")
        table["Tutar"] = table["Tutar"].apply(lambda x: f"{smart_to_num(x):,.2f} USD" if str(x).strip() else "")
        st.dataframe(table.drop(columns=["ID"]), use_container_width=True)

        option_labels = {"": "— Sipariş Seç —"}
        for _, row in pending_orders.iterrows():
            label = f"{row['Müşteri Adı']} - {row['Proforma No']}"
            termin_dt = pd.to_datetime(row.get("Termin Tarihi", ""), errors="coerce")
            if pd.notnull(termin_dt):
                label += f" | Termin: {termin_dt.strftime('%d/%m/%Y')}"
            tutar_raw = row.get("Tutar", "")
            if str(tutar_raw).strip():
                label += f" | Tutar: {smart_to_num(tutar_raw):,.2f} USD"
            option_labels[row["ID"]] = label

        pending_options = [""] + pending_orders["ID"].tolist()
        if st.session_state.get(pending_reset_flag_key):
            st.session_state[pending_select_key] = ""
            st.session_state[pending_reset_flag_key] = False

        if st.session_state[pending_select_key] not in pending_options:
            st.session_state[pending_select_key] = ""

        selected_pending = st.selectbox(
            "Fatura kaydı açmak istediğiniz siparişi seçin",
            options=pending_options,
            key=pending_select_key,
            format_func=lambda oid: option_labels.get(oid, "— Sipariş Seç —"),
        )

        if st.button("Seçimi Fatura Formuna Aktar", disabled=(selected_pending == "")):
            row = pending_orders[pending_orders["ID"] == selected_pending]
            if not row.empty:
                hedef = row.iloc[0]
                st.session_state[musteri_key] = str(hedef.get("Müşteri Adı", ""))
                st.session_state[proforma_key] = str(hedef.get("Proforma No", ""))
                st.session_state[pending_reset_flag_key] = True
                st.rerun()

   
    
    # ---- Müşteri / Proforma seçimleri ----
    st.markdown("### Fatura Ekle")
    
    musteri_secenek = sorted(df_proforma["Müşteri Adı"].dropna().astype(str).unique().tolist())
    musteri_options = [""] + musteri_secenek
    if st.session_state[musteri_key] not in musteri_options:
        st.session_state[musteri_key] = ""
    secilen_musteri = st.selectbox("Müşteri Seç", musteri_options, key=musteri_key)

    if st.session_state.get("invoice_last_customer") != secilen_musteri:
        st.session_state["invoice_last_customer"] = secilen_musteri
        st.session_state[proforma_key] = ""

    if secilen_musteri:
        p_list = (
            df_proforma.loc[
                df_proforma["Müşteri Adı"].astype(str) == secilen_musteri,
                "Proforma No",
            ]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        proforma_options = [""] + sorted(p_list)
    else:
        proforma_options = [""]

    if st.session_state[proforma_key] not in proforma_options:
        st.session_state[proforma_key] = ""

    proforma_no_sec = st.selectbox("Proforma No Seç", proforma_options, key=proforma_key)

    # ---- Müşteri varsayılanları (ülke/temsilci/ödeme) ----
    musteri_info = df_musteri[
        df_musteri["Müşteri Adı"].astype(str) == secilen_musteri
    ]
    ulke = musteri_info["Ülke"].values[0] if not musteri_info.empty else ""
    temsilci = musteri_info["Satış Temsilcisi"].values[0] if not musteri_info.empty else ""
    odeme = musteri_info["Ödeme Şekli"].values[0] if not musteri_info.empty else ""

    # ---- Proforma'dan Vade (gün) çek ve Vade Tarihi hesapla ----
    vade_gun = ""
    if secilen_musteri and proforma_no_sec:
        pr = df_proforma[
            (df_proforma["Müşteri Adı"].astype(str) == secilen_musteri)
            & (df_proforma["Proforma No"].astype(str) == proforma_no_sec)
        ]
        if not pr.empty:
            vade_gun = pr.iloc[0].get("Vade (gün)", "")

    # ---- Eski evrak linkleri (aynı müşteri+proforma altında son satır) ----
    onceki_evrak = df_evrak[
        (df_evrak["Müşteri Adı"].astype(str) == secilen_musteri)
        & (df_evrak["Proforma No"].astype(str) == proforma_no_sec)
    ].tail(1)

    def file_link_html(label, url):
        return f'<div style="margin-top:-6px;"><a href="{url}" target="_blank" style="color:#219A41;">[Daha önce yüklenmiş {label}]</a></div>' if url else \
               '<div style="margin-top:-6px; color:#b00020; font-size:0.95em;">(Daha önce yüklenmemiş)</div>'

    evrak_tipleri = [
        ("Commercial Invoice",  "Commercial Invoice PDF"),
        ("Sağlık Sertifikası",  "Sağlık Sertifikası PDF"),
        ("Packing List",        "Packing List PDF"),
        ("Konşimento",          "Konşimento PDF"),
        ("İhracat Beyannamesi", "İhracat Beyannamesi PDF"),
        ("Fatura PDF",          "Fatura PDF")  # eklendi
    ]

    # ---- Drive'a yükleme yardımcı fonksiyonu ----
    def upload_to_drive(parent_folder_id: str, filename: str, bytes_data: bytes) -> str:
        meta = {'title': filename, 'parents': [{'id': parent_folder_id}]}
        gfile = drive.CreateFile(meta)
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(bytes_data)
            tmp_path = tmp.name
        gfile.SetContentFile(tmp_path)
        try:
            gfile.Upload()  # personal drive: supportsAllDrives gerekmez
            link = f"https://drive.google.com/file/d/{gfile['id']}/view?usp=sharing"
        finally:
            try: os.remove(tmp_path)
            except: pass
        return link

    # ---- Form ----
    with st.form("add_evrak"):
        fatura_no = st.text_input("Fatura No")
        fatura_tarih = st.date_input("Fatura Tarihi", value=datetime.date.today())
        tutar = st.text_input("Fatura Tutarı (USD)")
        # Vade (gün) & vade tarihi gösterimi
        st.text_input("Vade (gün)", value=str(vade_gun), key="vade_gun", disabled=True)

        try:
            vade_int = int(vade_gun)
            vade_tarihi_hesap = fatura_tarih + datetime.timedelta(days=vade_int)
        except:
            vade_tarihi_hesap = None
        st.date_input("Vade Tarihi", value=(vade_tarihi_hesap or fatura_tarih), key="vade_tarihi", disabled=True)

        st.text_input("Ülke", value=ulke, disabled=True)
        st.text_input("Satış Temsilcisi", value=temsilci, disabled=True)
        st.text_input("Ödeme Şekli", value=odeme, disabled=True)

        # Evrak yüklemeleri + eski link gösterimleri
        uploaded_files = {}
        for col, label in evrak_tipleri:
            uploaded_files[col] = st.file_uploader(label, type="pdf", key=f"{col}_upload")
            prev_url = onceki_evrak.iloc[0][col] if not onceki_evrak.empty else ""
            st.markdown(file_link_html(label, prev_url), unsafe_allow_html=True)

        submitted = st.form_submit_button("Kaydet")

    if submitted:
        if not (secilen_musteri and proforma_no_sec and fatura_no.strip() and tutar.strip()):
            st.error("Müşteri, Proforma No, Fatura No ve Tutar zorunludur.")
            st.stop()

        tutar_num = smart_to_num(tutar)

        # 1) Dosyaları Drive'a yükle (varsa). Yoksa eski linki koru.
        file_urls = {}
        for col, _label in evrak_tipleri:
            upfile = uploaded_files[col]
            if upfile:
                clean_name = re.sub(r'[\\/*?:"<>|]+', "_", f"{secilen_musteri}__{proforma_no_sec}__{col}__{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
                file_urls[col] = upload_to_drive(EVRAK_KLASOR_ID, clean_name, upfile.read())
            else:
                file_urls[col] = onceki_evrak.iloc[0][col] if not onceki_evrak.empty else ""

        # 2) Tekilleştirme: aynı (Müşteri, Proforma, Fatura No) varsa GÜNCELLE; yoksa EKLE
        key_mask = (
            (df_evrak["Müşteri Adı"].astype(str) == secilen_musteri) &
            (df_evrak["Proforma No"].astype(str) == proforma_no_sec) &
            (df_evrak["Fatura No"].astype(str) == fatura_no)
        )

        # Vade Tarihi yazımı
        vade_tarihi_yaz = vade_tarihi_hesap if vade_tarihi_hesap else ""

        if key_mask.any():
            idx = df_evrak[key_mask].index[0]
            df_evrak.at[idx, "Fatura Tarihi"]    = fatura_tarih
            df_evrak.at[idx, "Tutar"]            = tutar
            df_evrak.at[idx, "Tutar_num"]        = tutar_num
            df_evrak.at[idx, "Vade (gün)"]       = vade_gun
            df_evrak.at[idx, "Vade Tarihi"]      = vade_tarihi_yaz
            df_evrak.at[idx, "Ülke"]             = ulke
            df_evrak.at[idx, "Satış Temsilcisi"] = temsilci
            df_evrak.at[idx, "Ödeme Şekli"]      = odeme
            mevcut_odeme = pd.to_numeric(
                pd.Series(df_evrak.at[idx, "Ödenen Tutar"]), errors="coerce"
            ).fillna(0.0).iloc[0]
            tutar_float = float(tutar_num) if pd.notnull(tutar_num) else 0.0
            df_evrak.at[idx, "Ödenen Tutar"] = min(max(mevcut_odeme, 0.0), tutar_float)
            if tutar_float > 0 and df_evrak.at[idx, "Ödenen Tutar"] >= tutar_float - 0.01:
                df_evrak.at[idx, "Ödendi"] = True   
            for col, _ in evrak_tipleri:
                df_evrak.at[idx, col] = file_urls.get(col, "")
            islem = "güncellendi"
        else:
            new_row = {
                "ID": str(uuid.uuid4()),
                "Müşteri Adı": secilen_musteri,
                "Proforma No": proforma_no_sec,
                "Fatura No": fatura_no,
                "Fatura Tarihi": fatura_tarih,
                "Tutar": tutar,
                "Tutar_num": tutar_num,
                "Vade (gün)": vade_gun,
                "Vade Tarihi": vade_tarihi_yaz,
                "Ülke": ulke,
                "Satış Temsilcisi": temsilci,
                "Ödeme Şekli": odeme,
                "Ödenen Tutar": 0.0,       
                "Ödendi": False,
                **{col: file_urls.get(col, "") for col, _ in evrak_tipleri},
                "Sipariş Formu": "",
                "Yük Resimleri": "",
                "EK Belgeler": "",
            }
            df_evrak = pd.concat([df_evrak, pd.DataFrame([new_row])], ignore_index=True)
            islem = "eklendi"
        if "Tutar_num" not in df_evrak.columns:
            df_evrak["Tutar_num"] = pd.NA
        df_evrak["Tutar_num"] = pd.to_numeric(df_evrak["Tutar_num"], errors="coerce")
        if "Ödenen Tutar" not in df_evrak.columns:
            df_evrak["Ödenen Tutar"] = 0.0
        df_evrak["Ödenen Tutar"] = pd.to_numeric(df_evrak["Ödenen Tutar"], errors="coerce").fillna(0.0)
        df_evrak["Ödenen Tutar"] = df_evrak["Ödenen Tutar"].clip(lower=0)
        df_evrak["Ödenen Tutar"] = np.minimum(
            df_evrak["Ödenen Tutar"],
            df_evrak["Tutar_num"].fillna(np.inf)
        )

        update_excel()
        st.success(f"Evrak {islem}!")
        st.rerun()

    st.markdown("### Kayıtlı Fatura Güncelle")

    invoice_mask = df_evrak["Fatura No"].astype(str).str.strip() != ""
    existing_invoices = df_evrak[invoice_mask].copy()

    if existing_invoices.empty:
        st.info("Güncellenebilecek kayıtlı fatura bulunmuyor.")
    else:
        def _safe_date(value, fallback=None):
            if isinstance(value, datetime.datetime):
                return value.date()
            if isinstance(value, datetime.date):
                return value
            try:
                ts = pd.to_datetime(value, errors="coerce")
            except Exception:
                ts = pd.NaT
            if pd.isna(ts):
                return fallback
            return ts.date()

        def _format_invoice(idx):
            row = existing_invoices.loc[idx]
            musteri = str(row.get("Müşteri Adı", "")).strip() or "Müşteri Yok"
            fatura_no = str(row.get("Fatura No", "")).strip() or "Numara Yok"
            proforma = str(row.get("Proforma No", "")).strip()
            invoice_date = _safe_date(row.get("Fatura Tarihi"))
            due_date = _safe_date(row.get("Vade Tarihi"))

            parts = [f"{musteri}", f"Fatura: {fatura_no}"]
            if proforma:
                parts.append(f"Proforma: {proforma}")
            if invoice_date:
                parts.append(f"Fatura Tarihi: {invoice_date.strftime('%d/%m/%Y')}")
            if due_date:
                parts.append(f"Vade Tarihi: {due_date.strftime('%d/%m/%Y')}")
            return " | ".join(parts)

        invoice_indices = existing_invoices.index.tolist()
        selected_invoice = st.selectbox(
            "Güncellemek istediğiniz faturayı seçin",
            options=invoice_indices,
            format_func=_format_invoice,
            key="invoice_edit_select",
        )

        if invoice_indices:
            secili_satir = existing_invoices.loc[selected_invoice]
            default_invoice_date = _safe_date(secili_satir.get("Fatura Tarihi"), fallback=datetime.date.today())
            default_due_date = _safe_date(secili_satir.get("Vade Tarihi"), fallback=default_invoice_date)

            mevcut_tutar = secili_satir.get("Tutar", "")
            if pd.isna(mevcut_tutar):
                mevcut_tutar = ""
            else:
                mevcut_tutar = str(mevcut_tutar)

            with st.form("update_invoice_details"):
                yeni_fatura_tarihi = st.date_input("Yeni Fatura Tarihi", value=default_invoice_date)
                yeni_vade_tarihi = st.date_input("Yeni Vade Tarihi", value=default_due_date)
                yeni_tutar = st.text_input("Yeni Fatura Tutarı (USD)", value=mevcut_tutar)
                update_submitted = st.form_submit_button("Bilgileri Güncelle")

            if update_submitted:
                df_evrak.at[selected_invoice, "Fatura Tarihi"] = yeni_fatura_tarihi
                df_evrak.at[selected_invoice, "Vade Tarihi"] = yeni_vade_tarihi
                df_evrak.at[selected_invoice, "Tutar"] = yeni_tutar

                yeni_tutar_num = smart_to_num(yeni_tutar)
                df_evrak.at[selected_invoice, "Tutar_num"] = yeni_tutar_num
                
                try:
                    gun_farki = (pd.Timestamp(yeni_vade_tarihi) - pd.Timestamp(yeni_fatura_tarihi)).days
                    df_evrak.at[selected_invoice, "Vade (gün)"] = str(gun_farki)
                except Exception:
                    df_evrak.at[selected_invoice, "Vade (gün)"] = ""


                if "Ödenen Tutar" in df_evrak.columns:
                    mevcut_odeme = pd.to_numeric(
                        pd.Series(df_evrak.at[selected_invoice, "Ödenen Tutar"]), errors="coerce"
                    ).fillna(0.0).iloc[0]
                    tutar_float = float(yeni_tutar_num) if pd.notnull(yeni_tutar_num) else 0.0
                    guncel_odeme = min(max(mevcut_odeme, 0.0), tutar_float)
                    df_evrak.at[selected_invoice, "Ödenen Tutar"] = guncel_odeme

                    if "Ödendi" in df_evrak.columns:
                        if tutar_float > 0:
                            df_evrak.at[selected_invoice, "Ödendi"] = guncel_odeme >= tutar_float - 0.01
                        else:
                            df_evrak.at[selected_invoice, "Ödendi"] = False
                elif "Ödendi" in df_evrak.columns:
                    df_evrak.at[selected_invoice, "Ödendi"] = False
                    
                update_excel()
                st.success("Fatura bilgileri güncellendi!")
                st.rerun()

    st.markdown("### Fatura Kaydı Sil")

    if df_evrak.empty:
        st.info("Silinecek fatura kaydı bulunmuyor.")
    else:
        delete_options = df_evrak.index.tolist()

        def _format_delete_option(idx):
            row = df_evrak.loc[idx]
            musteri = str(row.get("Müşteri Adı", "")).strip() or "Müşteri Yok"
            fatura_no = str(row.get("Fatura No", "")).strip() or "Numara Yok"
            proforma_no = str(row.get("Proforma No", "")).strip()

            invoice_date = row.get("Fatura Tarihi", "")
            invoice_date_str = ""
            if pd.notna(invoice_date):
                invoice_ts = pd.to_datetime(invoice_date, errors="coerce")
                if pd.notna(invoice_ts):
                    invoice_date_str = invoice_ts.strftime("%d/%m/%Y")

            tutar_raw = row.get("Tutar", "")
            tutar_str = ""
            if str(tutar_raw).strip():
                tutar_str = f"{smart_to_num(tutar_raw):,.2f} USD"

            parts = [f"{musteri}", f"Fatura: {fatura_no}"]
            if proforma_no:
                parts.append(f"Proforma: {proforma_no}")
            if invoice_date_str:
                parts.append(f"Tarih: {invoice_date_str}")
            if tutar_str:
                parts.append(f"Tutar: {tutar_str}")
            return " | ".join(parts)

        with st.form("delete_invoice_form"):
            silinecek_fatura = st.selectbox(
                "Silmek istediğiniz faturayı seçin",
                options=delete_options,
                format_func=_format_delete_option,
                key="invoice_delete_select",
            )
            confirm_delete = st.checkbox("Silme işlemini onaylıyorum")
            delete_submitted = st.form_submit_button("Seçili Faturayı Sil")

        if delete_submitted:
            if confirm_delete:
                df_evrak = df_evrak.drop(index=silinecek_fatura).reset_index(drop=True)
                update_excel()
                st.success("Seçilen fatura kaydı silindi.")
                st.rerun()
            else:
                st.warning("Silme işlemini onaylamak için kutucuğu işaretleyin.")

    


### ===========================
### --- TAHSİLAT PLANI MENÜSÜ ---
### ===========================

elif menu == "Tahsilat Planı":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Tahsilat Planı</h2>", unsafe_allow_html=True)

    # Gerekli kolonlar yoksa ekle
    for c in ["Müşteri Adı","Fatura No","Fatura Tarihi","Vade Tarihi","Tutar_num","Ödenen Tutar","Ülke","Satış Temsilcisi","Ödeme Şekli","Ödendi"]:
        if c not in df_evrak.columns:
            if c == "Ödendi":
                df_evrak[c] = False
            elif c == "Ödenen Tutar" or c == "Tutar_num":
                df_evrak[c] = 0.0
            else:
                df_evrak[c] = ""

    # Sadece vadesi olan kayıtlar
    vade_df = df_evrak.copy()
    vade_df["Tutar_num"] = pd.to_numeric(vade_df["Tutar_num"], errors="coerce").fillna(0.0)
    vade_df["Ödenen Tutar"] = pd.to_numeric(vade_df["Ödenen Tutar"], errors="coerce").fillna(0.0)    
    vade_df["Vade Tarihi"] = pd.to_datetime(vade_df["Vade Tarihi"], errors="coerce")
    vade_df = vade_df[vade_df["Vade Tarihi"].notna()]
    vade_df["Kalan Bakiye"] = (vade_df["Tutar_num"] - vade_df["Ödenen Tutar"]).clip(lower=0.0)    

    if vade_df.empty:
        st.info("Vade tarihi girilmiş kayıt bulunmuyor.")
    else:
        today = pd.Timestamp.today().normalize()
        vade_df["Kalan Gün"] = (vade_df["Vade Tarihi"] - today).dt.days

        # Ödenmemişler üzerinden özet kutucukları
        acik_mask = vade_df["Kalan Bakiye"] > 0.01
        acik = vade_df[acik_mask].copy()
        vadesi_gelmemis = acik[acik["Kalan Gün"] > 0]
        bugun = acik[acik["Kalan Gün"] == 0]
        gecikmis = acik[acik["Kalan Gün"] < 0]

        c1, c2, c3 = st.columns(3)
        c1.metric("Vadeleri Gelmeyen", f"{float(vadesi_gelmemis['Kalan Bakiye'].sum()):,.2f} USD", f"{len(vadesi_gelmemis)} Fatura")
        c2.metric("Bugün Vadesi",   f"{float(bugun['Kalan Bakiye'].sum()):,.2f} USD", f"{len(bugun)} Fatura")
        c3.metric("Gecikmiş Ödemeler",        f"{float(gecikmis['Kalan Bakiye'].sum()):,.2f} USD", f"{len(gecikmis)} Fatura")

        st.markdown("---")

        # Filtreler
        f1, f2, f3 = st.columns([1.4, 1.2, 1.2])
        ulke_f = f1.multiselect("Ülke", sorted([u for u in vade_df["Ülke"].dropna().unique() if str(u).strip()]))
        tem_f  = f2.multiselect("Satış Temsilcisi", sorted([t for t in vade_df["Satış Temsilcisi"].dropna().unique() if str(t).strip()]))
        durum_f = f3.selectbox("Ödeme Durumu", ["Ödenmemiş (varsayılan)", "Hepsi", "Sadece Ödenmiş"], index=0)

        view = vade_df.copy()
        if ulke_f:
            view = view[view["Ülke"].isin(ulke_f)]
        if tem_f:
            view = view[view["Satış Temsilcisi"].isin(tem_f)]
        if durum_f == "Ödenmemiş (varsayılan)":
            view = view[view["Kalan Bakiye"] > 0.01]
        elif durum_f == "Sadece Ödenmiş":
            view = view[view["Kalan Bakiye"] <= 0.01]

        # Düzenlenebilir tablo (Ödendi sütunu için)
        editor_view = view.reset_index(drop=False).rename(columns={"index": "_row"})
        editor = editor_view[[
            "_row",
            "Müşteri Adı",
            "Ülke",
            "Satış Temsilcisi",
            "Fatura No",
            "Fatura Tarihi",
            "Vade Tarihi",
            "Kalan Gün",
            "Tutar_num",
            "Ödenen Tutar",
            "Kalan Bakiye",
            "Ödendi",
        ]].copy()

        editor["__fatura_order"] = pd.to_datetime(editor_view["Fatura Tarihi"], errors="coerce")
        editor["__vade_order"] = pd.to_datetime(editor_view["Vade Tarihi"], errors="coerce")

        editor_display = editor.sort_values(["Kalan Gün", "__vade_order"])
        editor_display["Fatura Tarihi"] = editor_display["__fatura_order"].dt.strftime("%d/%m/%Y").fillna("")
        editor_display["Vade Tarihi"] = editor_display["__vade_order"].dt.strftime("%d/%m/%Y").fillna("")
        editor_display["Tutar_num"] = pd.to_numeric(editor_display["Tutar_num"], errors="coerce").fillna(0.0)
        editor_display["Ödenen Tutar"] = pd.to_numeric(editor_display["Ödenen Tutar"], errors="coerce").fillna(0.0)
        editor_display["Kalan Bakiye"] = pd.to_numeric(editor_display["Kalan Bakiye"], errors="coerce").fillna(0.0)
        editor_display["Ödendi"] = editor_display["Ödendi"].astype(bool)

        edited = st.data_editor(
            editor_display.drop(columns=["__fatura_order", "__vade_order"]),
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "_row": st.column_config.Column("Satır ID", disabled=True, width="small"),
                "Tutar_num": st.column_config.NumberColumn("Fatura Tutarı (USD)", format="%.2f", disabled=True),
                "Ödenen Tutar": st.column_config.NumberColumn("Ödenen Tutar (USD)", format="%.2f", disabled=True),
                "Kalan Bakiye": st.column_config.NumberColumn("Kalan Bakiye (USD)", format="%.2f", disabled=True),
                "Ödendi": st.column_config.CheckboxColumn("Ödendi", help="Ödemesi tamamlandıysa işaretleyin."),
            },
            disabled=[
                "_row",
                "Müşteri Adı",
                "Ülke",
                "Satış Temsilcisi",
                "Fatura No",
                "Fatura Tarihi",
                "Vade Tarihi",
                "Kalan Gün",
                "Tutar_num",
                "Ödenen Tutar",
                "Kalan Bakiye",
            ],
            key="tahsilat_editor",
        )

        if not edited.empty:
            edited["_row"] = pd.to_numeric(edited["_row"], errors="coerce").fillna(-1).astype(int)
            edited["Ödendi"] = edited["Ödendi"].astype(bool)
            original_flags = editor_view.set_index("_row")["Ödendi"].astype(bool)
            merged = edited.merge(
                original_flags.reset_index().rename(columns={"Ödendi": "_original_odendi"}),
                on="_row",
                how="left",
            )
            changed = merged[merged["Ödendi"] != merged["_original_odendi"]]

            if not changed.empty:
                for _, row in changed.iterrows():
                    idx = int(row["_row"])
                    new_flag = bool(row["Ödendi"])
                    df_evrak.at[idx, "Ödendi"] = new_flag

                    toplam = pd.to_numeric(pd.Series([df_evrak.at[idx, "Tutar_num"]]), errors="coerce").fillna(0.0).iloc[0]
                    mevcut = pd.to_numeric(pd.Series([df_evrak.at[idx, "Ödenen Tutar"]]), errors="coerce").fillna(0.0).iloc[0]

                    if new_flag:
                        df_evrak.at[idx, "Ödenen Tutar"] = round(max(toplam, 0.0), 2)
                    else:
                        df_evrak.at[idx, "Ödenen Tutar"] = round(min(mevcut, toplam), 2)

                update_excel()
                st.success(f"{len(changed)} kaydın ödeme durumu güncellendi.")
                st.rerun()

        st.markdown("#### Ödeme Durumu Güncelle")
        if not view.empty:
            # ID yoksa güvenli seçim için bir satır anahtarı oluşturalım
            view = view.reset_index(drop=False).rename(columns={"index":"_row"})
            sec = st.selectbox(
                "Kayıt Seç",
                options=view["_row"].tolist(),
                format_func=lambda i: f"{view.loc[view['_row']==i,'Müşteri Adı'].values[0]} | {view.loc[view['_row']==i,'Fatura No'].values[0]}"
            )

            secili = view.loc[view["_row"] == sec].iloc[0]
            toplam_tutar = float(secili.get("Tutar_num", 0.0) or 0.0)
            odenen_tutar = float(pd.to_numeric(pd.Series(secili.get("Ödenen Tutar", 0.0)), errors="coerce").fillna(0.0).iloc[0])
            if odenen_tutar < 0:
                odenen_tutar = 0.0
            kalan_bakiye = max(toplam_tutar - odenen_tutar, 0.0)

            ozet_cols = st.columns(3)
            ozet_cols[0].metric("Fatura Tutarı", f"{toplam_tutar:,.2f} USD")
            ozet_cols[1].metric("Ödenen Tutar", f"{odenen_tutar:,.2f} USD")
            ozet_cols[2].metric("Kalan Bakiye", f"{kalan_bakiye:,.2f} USD")

            min_ara_odeme = -odenen_tutar
            max_ara_odeme = max(toplam_tutar - odenen_tutar, 0.0)

            with st.form(f"tahsilat_guncelle_{sec}"):
                mevcut_odendi = bool(secili.get("Ödendi", False))
                ara_odeme = st.number_input(
                    "Ara ödeme tutarı (USD)",
                    min_value=float(min_ara_odeme),
                    max_value=float(max_ara_odeme),
                    value=0.0,
                    step=100.0,
                    format="%.2f",
                    key=f"ara_odeme_input_{sec}"
                )
                st.caption("Not: Gerekirse eksi tutar girerek önceki tahsilatı azaltabilirsiniz.")
                odendi_mi = st.checkbox(
                    "Ödendi olarak işaretle",
                    value=mevcut_odendi,
                    key=f"odendi_checkbox_{sec}"
                )
                kaydet = st.form_submit_button("Kaydet / Güncelle")

            if kaydet:
                ana_index = int(view.loc[view["_row"] == sec, "_row"].iloc[0])
                yeni_odenen = odenen_tutar + float(ara_odeme)
                yeni_odenen = max(min(yeni_odenen, toplam_tutar), 0.0)
                if toplam_tutar <= 0:
                    yeni_odenen = 0.0
                if odendi_mi or yeni_odenen >= max(toplam_tutar - 0.01, 0):
                    yeni_odenen = toplam_tutar
                    odendi_mi = True
                else:
                    odendi_mi = yeni_odenen >= max(toplam_tutar - 0.01, 0)
                df_evrak.at[ana_index, "Ödenen Tutar"] = round(yeni_odenen, 2)
                df_evrak.at[ana_index, "Ödendi"] = bool(odendi_mi)
                update_excel()
                st.success("Tahsilat bilgisi güncellendi!")
                st.rerun()

### ===========================
### --- ETA İZLEME MENÜSÜ ---
### ===========================

elif menu == "ETA İzleme":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>ETA İzleme</h2>", unsafe_allow_html=True)

    # ---- Sabitler ----
    ROOT_EXPORT_FOLDER_ID = "14FTE1oSeIeJ6Y_7C0oQyZPKC8dK8hr1J"  # İhracat Evrakları ana klasör ID (MY DRIVE)

    # ---- Güvenlik: gerekli kolonlar ----
    for col in ["Sevk Durumu", "Proforma No", "Sevk Tarihi", "Ulaşma Tarihi"]:
        if col not in df_proforma.columns:
            df_proforma[col] = ""
            
    for col in ETA_COLUMNS:            
        if col not in df_eta.columns:
            df_eta[col] = ""
    extra_eta_cols = [col for col in df_eta.columns if col not in ETA_COLUMNS]
    df_eta = df_eta.reindex(columns=ETA_COLUMNS + extra_eta_cols, fill_value="")            

    # ---- Yardımcılar ----
    def safe_name(text, maxlen=120):
        s = str(text or "").strip()
        s = re.sub(r"\s+", " ", s)            # çoklu boşluk -> tek
        s = s.replace(" ", "_")               # boşluk -> _
        s = re.sub(r'[\\/*?:"<>|]+', "_", s)  # Drive yasak karakterleri
        return s[:maxlen]

    def get_or_create_folder_by_name(name: str, parent_id: str) -> str:
        """
        Parent altında isme göre klasör bulur; yoksa oluşturur.
        (My Drive — Shared Drive kullanılmıyor)
        """
        q = (
            f"title = '{name}' and mimeType = 'application/vnd.google-apps.folder' "
            f"and '{parent_id}' in parents and trashed = false"
        )
        try:
            lst = drive.ListFile({'q': q}).GetList()
            if lst:
                return lst[0]['id']
            meta = {
                'title': name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [{'id': parent_id}],
            }
            f = drive.CreateFile(meta)
            f.Upload()
            return f['id']
        except Exception as e:
            st.error(f"Klasör oluşturma/arama hatası: {e}")
            return ""

    def resolve_folder_date(musteri: str, proforma_no: str) -> datetime.date:
        """
        Klasör adı için kullanılacak tarihi belirler:
        1) Proforma 'Sevk Tarihi' varsa o,
        2) yoksa ETA kaydındaki 'Sevk Tarihi',
        3) yoksa ilgili ETA kaydındaki 'ETA Tarihi',
        4) o da yoksa bugün.
        """
        # Sevk Tarihi
        pr_mask = (df_proforma["Müşteri Adı"] == musteri) & (df_proforma["Proforma No"] == proforma_no)
        sevk_ts = None
        if pr_mask.any():
            try:
                sevk_ts = pd.to_datetime(df_proforma.loc[pr_mask, "Sevk Tarihi"].values[0], errors="coerce")
            except Exception:
                sevk_ts = None
        if pd.notnull(sevk_ts):
            try:
                return sevk_ts.date()
            except Exception:
                pass

        # ETA Sevk Tarihi
        eta_mask = (df_eta["Müşteri Adı"] == musteri) & (df_eta["Proforma No"] == proforma_no)
        eta_sevk_ts = None
        if eta_mask.any():
            try:
                eta_sevk_ts = pd.to_datetime(df_eta.loc[eta_mask, "Sevk Tarihi"].values[0], errors="coerce")
            except Exception:
                eta_sevk_ts = None
        if pd.notnull(eta_sevk_ts):
            try:
                return eta_sevk_ts.date()
            except Exception:
                pass

        # ETA Tarihi        
        eta_ts = None
        if eta_mask.any():
            try:
                eta_ts = pd.to_datetime(df_eta.loc[eta_mask, "ETA Tarihi"].values[0], errors="coerce")
            except Exception:
                eta_ts = None
        if pd.notnull(eta_ts):
            try:
                return eta_ts.date()
            except Exception:
                pass

        # Default: bugün
        return datetime.date.today()

    def get_loading_photos_folder(musteri_adi: str, proforma_no: str, fallback_date: datetime.date) -> str:
        """
        Ana klasör altında <Müşteri_Adi>_<Proforma_No> / Yükleme Resimleri hiyerarşisini hazırlar ve döndürür.
        Proforma numarası yoksa tarih tabanlı bir isimlendirmeye geri döner.
        """
        if not ROOT_EXPORT_FOLDER_ID:
            return ""
  
        musteri_parca = safe_name(musteri_adi)
        proforma_parca = safe_name(proforma_no)
        if not proforma_parca:
            proforma_parca = fallback_date.strftime('%Y-%m-%d')

        folder_name = f"{musteri_parca}_{proforma_parca}" if musteri_parca else proforma_parca
        parent = get_or_create_folder_by_name(folder_name, ROOT_EXPORT_FOLDER_ID)
        if not parent:
            return ""
        yukleme = get_or_create_folder_by_name("Yükleme Resimleri", parent)
        return yukleme

    # ==== SEVKEDİLENLER (Yolda) ====
    sevkedilenler = df_proforma[df_proforma["Sevk Durumu"] == "Sevkedildi"].copy()
    if sevkedilenler.empty:
        st.info("Sevkedilmiş sipariş bulunmuyor.")
    else:
        # Seçim
        secenekler = sevkedilenler[["Müşteri Adı", "Proforma No"]].drop_duplicates()
        secenekler["sec_text"] = secenekler["Müşteri Adı"] + " - " + secenekler["Proforma No"]
        selected = st.selectbox("Sevkedilen Sipariş Seç", secenekler["sec_text"])
        selected_row = secenekler[secenekler["sec_text"] == selected].iloc[0]
        sec_musteri = selected_row["Müşteri Adı"]
        sec_proforma = selected_row["Proforma No"]

        # === Klasör bilgisi (Sevk/ETA/bugün) + Müşteri adı ===
        klasor_tarih = resolve_folder_date(sec_musteri, sec_proforma)
        proforma_gosterim = ""
        if pd.notna(sec_proforma):
            proforma_gosterim = str(sec_proforma).strip()
        if not proforma_gosterim:
            proforma_gosterim = klasor_tarih.strftime('%Y-%m-%d')
            
        # ========== YÜKLEME FOTOĞRAFLARI (Müşteri_Adi + Proforma → “Yükleme Resimleri”) ==========
        st.markdown("#### Yükleme Fotoğrafları (Müşteri + Proforma bazlı)")

        hedef_klasor = get_loading_photos_folder(sec_musteri, sec_proforma, klasor_tarih)
        if not hedef_klasor:
            st.error("Klasör hiyerarşisi oluşturulamadı.")
        else:
            # 1) Klasörü yeni sekmede aç butonu
            drive_link = f"https://drive.google.com/drive/folders/{hedef_klasor}?usp=sharing"
            st.markdown(f"[Klasörü yeni sekmede aç]({drive_link})")

            # 2) Panel içinde gömülü görüntüleme – sadece gezinme
            with st.expander(f"Panelde klasörü görüntüle – {sec_musteri} / {proforma_gosterim}"):
                embed = f"https://drive.google.com/embeddedfolderview?id={hedef_klasor}#grid"
                st.markdown(
                    f'<iframe src="{embed}" width="100%" height="520" frameborder="0" '
                    f'style="border:1px solid #eee; border-radius:12px;"></iframe>',
                    unsafe_allow_html=True
                )

            # 3) Mevcut dosyaları say ve özetle (ilk 10 isim)
            try:
                mevcut_dosyalar = drive.ListFile({
                    'q': f"'{hedef_klasor}' in parents and trashed = false"
                }).GetList()
            except Exception as e:
                mevcut_dosyalar = []
                st.warning(f"Dosyalar listelenemedi: {e}")

            if mevcut_dosyalar:
                st.caption(f"Bu klasörde {len(mevcut_dosyalar)} dosya var.")
                names = [f"- {f['title']}" for f in mevcut_dosyalar[:10]]
                st.write("\n".join(names) if names else "")
                if len(mevcut_dosyalar) > 10:
                    st.write("…")

            # 4) (OPSİYONEL) Dosya Ekle – duplike önleme (aynı isim SKIP)
            with st.expander("Dosya Ekle (opsiyonel, duplike önleme)"):
                files = st.file_uploader(
                    "Yüklenecek dosyaları seçin",
                    type=["pdf", "jpg", "jpeg", "png", "webp"],
                    accept_multiple_files=True,
                    key=f"yuk_resimleri_dedupe_{safe_name(sec_musteri)}_{safe_name(proforma_gosterim)}"
                )

                if files:
                    var_olan_isimler = set(f["title"] for f in mevcut_dosyalar)
                    yuklenen_say = 0
                    atlanan_duplike = 0

                    for up in files:
                        suffix = os.path.splitext(up.name)[1].lower() or ""
                        base = os.path.splitext(up.name)[0]
                        fname = safe_name(base) + suffix

                        if fname in var_olan_isimler:
                            atlanan_duplike += 1
                            continue

                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as fp:
                            fp.write(up.read())
                            temp_path = fp.name

                        meta = {'title': fname, 'parents': [{'id': hedef_klasor}]}
                        gfile = drive.CreateFile(meta)
                        gfile.SetContentFile(temp_path)
                        try:
                            gfile.Upload()
                            yuklenen_say += 1
                            var_olan_isimler.add(fname)
                        except Exception as e:
                            st.error(f"{up.name} yüklenemedi: {e}")
                        finally:
                            try: os.remove(temp_path)
                            except: pass

                    if yuklenen_say:
                        update_excel()
                        st.success(f"{yuklenen_say} yeni dosya yüklendi.")
                        if atlanan_duplike:
                            st.info(f"{atlanan_duplike} dosya aynı isimle bulunduğu için atlandı.")
                        st.rerun()
                    else:
                        if atlanan_duplike and not yuklenen_say:
                            st.warning("Tüm dosyalar klasörde zaten mevcut (isimler aynı).")

        st.markdown("---")

        # ========== ETA Düzenleme ==========
        # Önceden ETA girilmiş mi?
        filtre = (df_eta["Müşteri Adı"] == sec_musteri) & (df_eta["Proforma No"] == sec_proforma)
        if filtre.any():
            mevcut_eta = df_eta.loc[filtre, "ETA Tarihi"].values[0]
            mevcut_aciklama = df_eta.loc[filtre, "Açıklama"].values[0],
            mevcut_sevk = df_eta.loc[filtre, "Sevk Tarihi"].values[0]            
        else:
            mevcut_eta = ""
            mevcut_aciklama = ""
            mevcut_sevk = ""
        proforma_mask = (df_proforma["Müşteri Adı"] == sec_musteri) & (df_proforma["Proforma No"] == sec_proforma)
        mevcut_proforma_sevk = df_proforma.loc[proforma_mask, "Sevk Tarihi"].values[0] if proforma_mask.any() else ""

        def _safe_date(value):
            if value is None:
                return None
            if isinstance(value, str) and not value.strip():
                return None
            try:
                ts = pd.to_datetime(value, errors="coerce")
                
            except Exception:
                return None
            if pd.isna(ts):
                return None
            try:
                return ts.date()
            except Exception:
                return None

        with st.form("edit_eta"):
            varsayilan_eta = _safe_date(mevcut_eta) or datetime.date.today()
            varsayilan_sevk = _safe_date(mevcut_sevk) or _safe_date(mevcut_proforma_sevk) or datetime.date.today()
            sevk_tarih = st.date_input("Sevk Tarihi", value=varsayilan_sevk)          
            eta_tarih = st.date_input("ETA Tarihi", value=varsayilan_eta)
            aciklama = st.text_area("Açıklama", value=mevcut_aciklama)
            guncelle = st.form_submit_button("ETA'yı Kaydet/Güncelle")
            ulasti = st.form_submit_button("Ulaştı")
            geri_al = st.form_submit_button("Sevki Geri Al")

            if guncelle:
                if filtre.any():
                    df_eta.loc[filtre, "Sevk Tarihi"] = sevk_tarih                    
                    df_eta.loc[filtre, "ETA Tarihi"] = eta_tarih
                    df_eta.loc[filtre, "Açıklama"] = aciklama
                else:
                    new_row = {
                        "Müşteri Adı": sec_musteri,
                        "Proforma No": sec_proforma,
                        "Sevk Tarihi": sevk_tarih,                        
                        "ETA Tarihi": eta_tarih,
                        "Açıklama": aciklama
                    }
                    df_eta = pd.concat([df_eta, pd.DataFrame([new_row])], ignore_index=True)
                if proforma_mask.any():
                    df_proforma.loc[proforma_mask, "Sevk Tarihi"] = sevk_tarih                 
                update_excel()
                st.success("ETA kaydedildi/güncellendi!")
                st.rerun()

            if ulasti:
                # Ulaşıldı: ETA listesinden çıkar, proforma'da Sevk Durumu "Ulaşıldı" ve bugünün tarihi "Ulaşma Tarihi" olarak kaydet
                df_eta = df_eta[~((df_eta["Müşteri Adı"] == sec_musteri) & (df_eta["Proforma No"] == sec_proforma))]
                idx = df_proforma[(df_proforma["Müşteri Adı"] == sec_musteri) & (df_proforma["Proforma No"] == sec_proforma)].index
                if len(idx) > 0:
                    df_proforma.at[idx[0], "Sevk Durumu"] = "Ulaşıldı"
                    df_proforma.at[idx[0], "Ulaşma Tarihi"] = datetime.date.today()
                update_excel()
                st.success("Sipariş 'Ulaşıldı' olarak işaretlendi ve ETA takibinden çıkarıldı!")
                st.rerun()

            if geri_al:
                # Siparişi geri al: ETA'dan çıkar, proforma'da sevk durumunu boş yap (Sipariş Operasyonları'na döner)
                df_eta = df_eta[~((df_eta["Müşteri Adı"] == sec_musteri) & (df_eta["Proforma No"] == sec_proforma))]
                idx = df_proforma[(df_proforma["Müşteri Adı"] == sec_musteri) & (df_proforma["Proforma No"] == sec_proforma)].index
                if len(idx) > 0:
                    df_proforma.at[idx[0], "Sevk Durumu"] = ""
                update_excel()
                st.success("Sevkiyat geri alındı! Sipariş tekrar Sipariş Operasyonları'na gönderildi.")
                st.rerun()

    # ==== ETA TAKİP LİSTESİ ====
    st.markdown("#### ETA Takip Listesi")
    for col in ["Proforma No", "Sevk Tarihi", "ETA Tarihi"]:
        if col not in df_eta.columns:
            df_eta[col] = ""
    if not df_eta.empty:
        df_eta_display = df_eta.copy()
        df_eta_display["ETA Tarihi"] = pd.to_datetime(
            df_eta_display["ETA Tarihi"], errors="coerce", dayfirst=True
        )
        df_eta_display["Sevk Tarihi"] = pd.to_datetime(
            df_eta_display["Sevk Tarihi"], errors="coerce", dayfirst=True
        )
        df_eta_display["ETA Tarihi"] = df_eta_display["ETA Tarihi"].dt.normalize()
        df_eta_display["Sevk Tarihi"] = df_eta_display["Sevk Tarihi"].dt.normalize()
        today = pd.Timestamp.today().normalize()
        df_eta_display["Kalan Gün"] = (df_eta_display["ETA Tarihi"] - today).dt.days
        df_eta_display = df_eta_display.sort_values(["ETA Tarihi", "Müşteri Adı", "Proforma No"], ascending=[True, True, True])
        tablo = df_eta_display[["Müşteri Adı", "Proforma No", "Sevk Tarihi", "ETA Tarihi", "Kalan Gün", "Açıklama"]].copy()
        tablo["ETA Tarihi"] = tablo["ETA Tarihi"].dt.strftime("%d/%m/%Y")
        tablo["Sevk Tarihi"] = tablo["Sevk Tarihi"].dt.strftime("%d/%m/%Y")
        tablo["ETA Tarihi"] = tablo["ETA Tarihi"].fillna("").replace({"NaT": ""})
        tablo["Sevk Tarihi"] = tablo["Sevk Tarihi"].fillna("").replace({"NaT": ""})
        st.dataframe(tablo, use_container_width=True)

        st.markdown("##### ETA Kaydı Sil")
        silinecekler = df_eta.index.tolist()
        sil_sec = st.selectbox("Silinecek Kaydı Seçin", options=silinecekler,
            format_func=lambda i: f"{df_eta.at[i, 'Müşteri Adı']} - {df_eta.at[i, 'Proforma No']}")
        if st.button("KAYDI SİL"):
            df_eta = df_eta.drop(sil_sec).reset_index(drop=True)
            update_excel()
            st.success("Seçilen ETA kaydı silindi!")
            st.rerun()
    else:
        st.info("Henüz ETA kaydı yok.")

    # ==== ULAŞANLAR (TESLİM EDİLENLER) ====
    ulasanlar = df_proforma[df_proforma["Sevk Durumu"] == "Ulaşıldı"].copy()

    if not ulasanlar.empty:
        ulasanlar["sec_text"] = ulasanlar["Müşteri Adı"] + " - " + ulasanlar["Proforma No"]
        st.markdown("#### Teslim Edilen Siparişlerde İşlemler")
        selected_ulasan = st.selectbox("Sipariş Seçiniz", ulasanlar["sec_text"])
        row = ulasanlar[ulasanlar["sec_text"] == selected_ulasan].iloc[0]

        # Ulaşma tarihi düzenleme
        try:
            current_ulasma = pd.to_datetime(row.get("Ulaşma Tarihi", None)).date()
            if pd.isnull(current_ulasma) or str(current_ulasma) == "NaT":
                current_ulasma = datetime.date.today()
        except Exception:
            current_ulasma = datetime.date.today()

        new_ulasma_tarih = st.date_input("Ulaşma Tarihi", value=current_ulasma, key="ulasan_guncelle")
        if st.button("Ulaşma Tarihini Kaydet"):
            idx = df_proforma[(df_proforma["Müşteri Adı"] == row["Müşteri Adı"]) & 
                              (df_proforma["Proforma No"] == row["Proforma No"])].index
            if len(idx) > 0:
                df_proforma.at[idx[0], "Ulaşma Tarihi"] = new_ulasma_tarih
                update_excel()
                st.success("Ulaşma Tarihi güncellendi!")
                st.rerun()

        st.markdown("---")
        # Ulaşanlardan YOLA GERİ AL (yeniden Sevkedildi + ETA’ya ekle/güncelle)
        with st.form("ulasan_geri_al_form"):
            st.markdown("##### 🔄 Ulaşan siparişi yeniden **Yolda Olanlar (ETA)** listesine al")
            yeni_eta = st.date_input("Yeni ETA (opsiyonel)", value=datetime.date.today() + datetime.timedelta(days=7))
            aciklama_geri = st.text_input("Açıklama (opsiyonel)", value="Geri alındı - tekrar yolda")
            onay = st.form_submit_button("Yola Geri Al")

        if onay:
            musteri = row["Müşteri Adı"]
            pno = row["Proforma No"]

            # Proforma statüsü
            idx = df_proforma[(df_proforma["Müşteri Adı"] == musteri) & (df_proforma["Proforma No"] == pno)].index
            if len(idx) > 0:
                df_proforma.at[idx[0], "Sevk Durumu"] = "Sevkedildi"
                df_proforma.at[idx[0], "Ulaşma Tarihi"] = ""

            # ETA ekle/güncelle
            filtre_eta = (df_eta["Müşteri Adı"] == musteri) & (df_eta["Proforma No"] == pno)
            eta_deger = pd.to_datetime(yeni_eta) if yeni_eta else ""
            sevk_kaydi = df_proforma.at[idx[0], "Sevk Tarihi"] if len(idx) > 0 and "Sevk Tarihi" in df_proforma.columns else ""            
            if filtre_eta.any():
                if yeni_eta:
                    df_eta.loc[filtre_eta, "ETA Tarihi"] = eta_deger
                if aciklama_geri:
                    df_eta.loc[filtre_eta, "Açıklama"] = aciklama_geri
                df_eta.loc[filtre_eta, "Sevk Tarihi"] = sevk_kaydi                    
            else:
                yeni_satir = {
                    "Müşteri Adı": musteri,
                    "Proforma No": pno,
                    "Sevk Tarihi": sevk_kaydi,                    
                    "ETA Tarihi": eta_deger if yeni_eta else "",
                    "Açıklama": aciklama_geri,
                }
                df_eta = pd.concat([df_eta, pd.DataFrame([yeni_satir])], ignore_index=True)

            update_excel()
            st.success("Sipariş, Ulaşanlar'dan geri alındı ve ETA listesine taşındı (Sevkedildi).")
            st.rerun()

        # Ulaşanlar Tablosu
        st.markdown("#### Ulaşan (Teslim Edilmiş) Siparişler")
        if "Sevk Tarihi" in ulasanlar.columns:
            ulasanlar["Sevk Tarihi"] = pd.to_datetime(ulasanlar["Sevk Tarihi"], errors="coerce")
        else:
            ulasanlar["Sevk Tarihi"] = pd.NaT
        if "Termin Tarihi" in ulasanlar.columns:
            ulasanlar["Termin Tarihi"] = pd.to_datetime(ulasanlar["Termin Tarihi"], errors="coerce")
        else:
            ulasanlar["Termin Tarihi"] = pd.NaT
        if "Tarih" in ulasanlar.columns:
            ulasanlar["Proforma Tarihi"] = pd.to_datetime(ulasanlar["Tarih"], errors="coerce")
        else:
            ulasanlar["Proforma Tarihi"] = pd.NaT            
        ulasanlar["Ulaşma Tarihi"] = pd.to_datetime(ulasanlar["Ulaşma Tarihi"], errors="coerce")

        ulasanlar["Gün Farkı"] = (ulasanlar["Sevk Tarihi"] - ulasanlar["Proforma Tarihi"]).dt.days
        ulasanlar["Proforma Tarihi"] = ulasanlar["Proforma Tarihi"].dt.strftime("%d/%m/%Y")
        ulasanlar["Sevk Tarihi"] = ulasanlar["Sevk Tarihi"].dt.strftime("%d/%m/%Y")
        ulasanlar["Termin Tarihi"] = ulasanlar["Termin Tarihi"].dt.strftime("%d/%m/%Y")
        ulasanlar["Ulaşma Tarihi"] = ulasanlar["Ulaşma Tarihi"].dt.strftime("%d/%m/%Y")
        
        tablo = ulasanlar[["Müşteri Adı", "Proforma No", "Proforma Tarihi", "Termin Tarihi", "Sevk Tarihi", "Ulaşma Tarihi", "Gün Farkı", "Tutar", "Açıklama"]]
        st.dataframe(tablo, use_container_width=True)
    else:
        st.info("Henüz ulaşan sipariş yok.")

 

# ==============================
# FUAR KAYITLARI MENÜSÜ
# ==============================

# Gerekli kolonlar (eksikse ekle)
FUAR_KOLONLAR = [
    "Fuar Adı", "Müşteri Adı", "Ülke", "Telefon", "E-mail",
    "Satış Temsilcisi", "Açıklamalar", "Görüşme Kalitesi", "Tarih"
]
for c in FUAR_KOLONLAR:
    if c not in df_fuar_musteri.columns:
        df_fuar_musteri[c] = "" if c not in ["Görüşme Kalitesi", "Tarih"] else np.nan

if menu == "Fuar Kayıtları":
    st.markdown("<h2 style='color:#8e54e9; font-weight:bold; text-align:center;'>Fuar Kayıtları</h2>", unsafe_allow_html=True)
    st.info("Fuarlarda müşteri görüşmelerinizi hızlıca buraya ekleyin. Yeni kayıt oluşturun, mevcutları düzenleyin.")

    # --- Fuar seçimi / oluşturma ---
    mevcut_fuarlar = sorted([f for f in df_fuar_musteri["Fuar Adı"].dropna().unique() if str(f).strip() != ""])
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        fuar_secim_opsiyonlari = ["— Fuar Seçiniz —"] + mevcut_fuarlar
        varsayilan_indeks = 0
        onceki_fuar = st.session_state.get("_selected_fuar", "")
        if onceki_fuar:
            try:
                varsayilan_indeks = fuar_secim_opsiyonlari.index(onceki_fuar)
            except ValueError:
                varsayilan_indeks = 0
        fuar_secimi = st.selectbox("Fuar Seçiniz", fuar_secim_opsiyonlari, index=varsayilan_indeks)
        fuar_adi = "" if fuar_secimi == "— Fuar Seçiniz —" else fuar_secimi
        st.session_state["_selected_fuar"] = fuar_adi
    with col_f2:
        yeni_fuar = st.text_input("Yeni Fuar Adı (opsiyonel)")
        if st.button("Fuar Ekle"):
            y = yeni_fuar.strip()
            if not y:
                st.warning("Fuar adı boş olamaz.")
            elif y in mevcut_fuarlar:
                st.info("Bu fuar zaten mevcut.")
                st.session_state["_selected_fuar"] = y
            else:
                yeni_satir = {col: "" for col in df_fuar_musteri.columns}
                yeni_satir["Fuar Adı"] = y
                if "Görüşme Kalitesi" in yeni_satir:
                    yeni_satir["Görüşme Kalitesi"] = np.nan
                if "Tarih" in yeni_satir:
                    yeni_satir["Tarih"] = pd.NaT
                df_fuar_musteri = pd.concat(
                    [df_fuar_musteri, pd.DataFrame([yeni_satir])],
                    ignore_index=True,
                )
                update_excel()
                st.session_state["_selected_fuar"] = y
                st.success(f"Fuar eklendi: {y}")
                st.rerun()
                
    secim = st.radio("İşlem Seçiniz:", ["Yeni Kayıt", "Eski Kayıt"], horizontal=True)

    # --- YENİ KAYIT ---
    if secim == "Yeni Kayıt":
        st.markdown("#### Yeni Fuar Müşteri Kaydı")
        with st.form("fuar_musteri_ekle"):
            musteri_adi = st.text_input("Müşteri Adı")
            ulke = st.selectbox("Ülke Seçin", ulke_listesi)  # global listeden
            tel = st.text_input("Telefon")
            email = st.text_input("E-mail")
            temsilci = st.selectbox("Satış Temsilcisi", temsilci_listesi)  # global listeden
            aciklama = st.text_area("Açıklamalar")
            gorusme_kalitesi = st.slider("Görüşme Kalitesi (1=Kötü, 5=Çok İyi)", 1, 5, 3)
            tarih = st.date_input("Tarih", value=datetime.date.today())

            kaydet = st.form_submit_button("Kaydet")
            if kaydet:
                if not fuar_adi:
                    st.warning("Lütfen bir fuar seçin veya ekleyin.")
                elif not musteri_adi.strip():
                    st.warning("Müşteri adı gerekli.")
                else:
                    yeni = {
                        "Fuar Adı": fuar_adi,
                        "Müşteri Adı": musteri_adi.strip(),
                        "Ülke": ulke,
                        "Telefon": tel.strip(),
                        "E-mail": email.strip(),
                        "Satış Temsilcisi": temsilci,
                        "Açıklamalar": aciklama.strip(),
                        "Görüşme Kalitesi": int(gorusme_kalitesi),
                        "Tarih": tarih,
                    }
                    placeholder_indeksleri = df_fuar_musteri[
                        (df_fuar_musteri["Fuar Adı"] == fuar_adi)
                        & df_fuar_musteri["Müşteri Adı"].fillna("").astype(str).str.strip().eq("")
                    ].index

                    if len(placeholder_indeksleri) > 0:
                        hedef_index = placeholder_indeksleri[0]
                        for anahtar, deger in yeni.items():
                            df_fuar_musteri.at[hedef_index, anahtar] = deger
                    else:
                        df_fuar_musteri = pd.concat([df_fuar_musteri, pd.DataFrame([yeni])], ignore_index=True)
                    update_excel()
                    st.session_state["_selected_fuar"] = fuar_adi                    
                    st.success("Fuar müşterisi eklendi!")
                    st.rerun()

    # --- ESKİ KAYIT: listele / filtrele / düzenle / sil ---
    elif secim == "Eski Kayıt":
        if not fuar_adi:
            st.info("Önce bir fuar seçin.")
        else:
            st.markdown(f"<h4 style='color:#4776e6;'>{fuar_adi} – Kayıtlar</h4>", unsafe_allow_html=True)

            fuar_df = df_fuar_musteri[df_fuar_musteri["Fuar Adı"] == fuar_adi].copy()

            with st.expander("Toplu E-posta Gönderimi", expanded=False):
                email_list = extract_unique_emails(fuar_df.get("E-mail"))
                if not email_list:
                    st.info("Bu fuara ait kayıtlı e-posta adresi bulunamadı.")
                else:
                    multiselect_options = ["Tümünü seç"] + email_list
                    selected_options = st.multiselect(
                        "E-posta Adresleri",
                        multiselect_options,
                        key=f"bulk_mail_recipients_{fuar_adi}"
                    )
                    
                    attachments_key = f"bulk_mail_files_{fuar_adi}"
                    attachments = st.session_state.get(attachments_key)

                    image_previews = []
                    inline_cid_map = {}
                    if attachments:
                        image_cid_state = st.session_state.setdefault("bulk_mail_image_cids", {})
                        for uploaded_file in attachments:
                            if uploaded_file is None:
                                continue
                            mime_type = uploaded_file.type or ""
                            if not mime_type.startswith("image/"):
                                continue
                            try:
                                file_bytes = uploaded_file.getvalue()
                            except Exception as exc:
                                st.error(f"{uploaded_file.name} okunurken hata oluştu: {exc}")
                                continue
                            attachment_key = f"{uploaded_file.name}:{len(file_bytes)}"
                            cid = image_cid_state.get(attachment_key)
                            if not cid:
                                cid = make_msgid()
                                image_cid_state[attachment_key] = cid
                            image_previews.append((uploaded_file.name, cid, attachment_key))
                            inline_cid_map[attachment_key] = cid

                    if image_previews:
                        st.markdown("**HTML gövdesine eklenecek görseller:**")
                        for file_name, cid, _ in image_previews:
                            cid_value = cid.strip("<>")
                            
                            st.code(
                                f'<img src="cid:{cid_value}" alt="{file_name}">',
                                language="html"
                            )                            
                 
                    if "Tümünü seç" in selected_options:
                        selected_recipients = email_list
                    else:
                        selected_recipients = selected_options

                   
                    
                    available_countries = sorted({
                        str(country).strip()
                        for country in fuar_df.get("Ülke", [])
                        if str(country).strip()
                    })
                    country_options = ["— Ülke Seçiniz —"] + available_countries
                    selected_country = st.selectbox(
                        "Ülke Seçiniz",
                        country_options,
                        key=f"bulk_mail_country_{fuar_adi}"
                    )
                    selected_country = (
                        "" if selected_country == "— Ülke Seçiniz —" else selected_country
                    )

                    subject_key = f"bulk_mail_subject_{fuar_adi}"
                    body_key = f"bulk_mail_body_{fuar_adi}"
                    template_key = f"{subject_key}_last_template"
                    language_key = f"{subject_key}_language"

                    st.session_state.setdefault(subject_key, "")
                    st.session_state.setdefault(body_key, "")

                    prev_language = st.session_state.get(language_key)
                    last_template = st.session_state.get(template_key)

                    language = COUNTRY_LANGUAGE_MAP.get(selected_country) if selected_country else None
                    template = FAIR_MAIL_TEMPLATES.get(language) if language else None

                    if prev_language != language and template:
                        last_template_subject = ""
                        last_template_body = ""
                        if isinstance(last_template, tuple) and len(last_template) == 3:
                            last_template_subject = last_template[1] or ""
                            last_template_body = last_template[2] or ""

                        current_subject = st.session_state.get(subject_key, "")
                        current_body = st.session_state.get(body_key, "")

                        if (not current_subject) or (current_subject == last_template_subject):
                            st.session_state[subject_key] = template.get("subject", "")

                        if (not current_body) or (current_body == last_template_body):
                            st.session_state[body_key] = template.get("body", "")

                        st.session_state[template_key] = (
                            language,
                            template.get("subject", ""),
                            template.get("body", ""),
                        )

                    if prev_language != language:
                        st.session_state[language_key] = language

                    subject = st.text_input("Konu", key=subject_key)
                    body = st.text_area("E-posta İçeriği", key=body_key)

                    st.file_uploader(
                        "Ek Dosyalar",
                        accept_multiple_files=True,
                        key=attachments_key,
                    )

                    attachments = st.session_state.get(attachments_key)
                    
                    if st.button("Gönder", key=f"bulk_mail_send_{fuar_adi}"):
                        if not selected_recipients:
                            st.warning("Lütfen en az bir e-posta adresi seçin.")
                        elif not subject.strip():
                            st.warning("Lütfen e-posta konusu girin.")
                        elif not body.strip():
                            st.warning("Lütfen e-posta içeriği girin.")
                        else:
                            
                            try:
                                send_fair_bulk_email(
                                    selected_recipients,
                                    subject.strip(),
                                    body,
                                    attachments or [],
                                    embed_images=EMBED_IMAGES,
                                    inline_cid_map=inline_cid_map
                                )
                        
                                st.success("E-postalar başarıyla gönderildi.")
                            except Exception as exc:
                                st.error(f"E-posta gönderilirken hata oluştu: {exc}")
            

            # Hızlı filtreler
            col_fa, col_fb, col_fc = st.columns([1, 1, 1])
            with col_fa:
                min_puan = st.slider("Min. Görüşme Kalitesi", 1, 5, 1)
            with col_fb:
                tarih_bas = st.date_input("Başlangıç Tarihi", value=datetime.date.today() - datetime.timedelta(days=30))
            with col_fc:
                tarih_bit = st.date_input("Bitiş Tarihi", value=datetime.date.today())

            # Tip dönüşümleri ve filtre uygula
            fuar_df["Görüşme Kalitesi"] = pd.to_numeric(fuar_df["Görüşme Kalitesi"], errors="coerce")
            fuar_df["Tarih"] = pd.to_datetime(fuar_df["Tarih"], errors="coerce").dt.normalize()

            tarih_bas_ts = pd.Timestamp(tarih_bas).normalize()
            tarih_bit_ts = pd.Timestamp(tarih_bit).normalize()

            mask = (
                (fuar_df["Görüşme Kalitesi"].fillna(0) >= min_puan) &
                (fuar_df["Tarih"] >= tarih_bas_ts) &
                (fuar_df["Tarih"] <= tarih_bit_ts)
            )
            fuar_df = fuar_df[mask].copy().sort_values("Tarih", ascending=False)

            if fuar_df.empty:
                st.info("Filtrelere uyan kayıt yok.")
            else:
                # Seçim
                secili_index = st.selectbox(
                    "Düzenlemek/Silmek istediğiniz kaydı seçin:",
                    fuar_df.index,
                    format_func=lambda i: f"{fuar_df.at[i, 'Müşteri Adı']} ({fuar_df.at[i, 'Tarih'].date() if pd.notnull(fuar_df.at[i, 'Tarih']) else ''})"
                )

                # Detay formu
                with st.form("kayit_duzenle"):
                    musteri_adi = st.text_input("Müşteri Adı", value=str(fuar_df.at[secili_index, "Müşteri Adı"]))
                    u_val = fuar_df.at[secili_index, "Ülke"]
                    ulke = st.selectbox("Ülke", ulke_listesi, index=ulke_listesi.index(u_val) if u_val in ulke_listesi else ulke_listesi.index("Diğer"))
                    t_val = fuar_df.at[secili_index, "Satış Temsilcisi"]
                    temsilci = st.selectbox("Satış Temsilcisi", temsilci_listesi, index=temsilci_listesi.index(t_val) if t_val in temsilci_listesi else 0)
                    tel = st.text_input("Telefon", value=str(fuar_df.at[secili_index, "Telefon"] or ""))
                    email = st.text_input("E-mail", value=str(fuar_df.at[secili_index, "E-mail"] or ""))
                    aciklama = st.text_area("Açıklamalar", value=str(fuar_df.at[secili_index, "Açıklamalar"] or ""))
                    gk_raw = fuar_df.at[secili_index, "Görüşme Kalitesi"]
                    gk_default = int(gk_raw) if pd.notnull(gk_raw) and str(gk_raw).isdigit() else 3
                    gorusme_kalitesi = st.slider("Görüşme Kalitesi (1-5)", 1, 5, gk_default)
                    t_raw = fuar_df.at[secili_index, "Tarih"]
                    tarih = st.date_input("Tarih", value=(pd.to_datetime(t_raw).date() if pd.notnull(t_raw) else datetime.date.today()))

                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        guncelle = st.form_submit_button("Kaydı Güncelle")
                    with col_b2:
                        sil = st.form_submit_button("Kaydı Sil")

                # Güncelle
                if guncelle:
                    for k, v in {
                        "Müşteri Adı": musteri_adi.strip(),
                        "Ülke": ulke,
                        "Telefon": tel.strip(),
                        "E-mail": email.strip(),
                        "Satış Temsilcisi": temsilci,
                        "Açıklamalar": aciklama.strip(),
                        "Görüşme Kalitesi": int(gorusme_kalitesi),
                        "Tarih": tarih,
                    }.items():
                        df_fuar_musteri.at[secili_index, k] = v
                    update_excel()
                    st.success("Kayıt güncellendi!")
                    st.rerun()

                # Sil
                if sil:
                    df_fuar_musteri = df_fuar_musteri.drop(secili_index).reset_index(drop=True)
                    update_excel()
                    st.success("Kayıt silindi!")
                    st.rerun()

                # Görsel tablo
                tablo = fuar_df.copy()
                tablo["Tarih"] = tablo["Tarih"].dt.strftime("%d/%m/%Y")
                st.dataframe(tablo[[
                    "Müşteri Adı", "Ülke", "Telefon", "E-mail",
                    "Satış Temsilcisi", "Açıklamalar", "Görüşme Kalitesi", "Tarih"
                ]], use_container_width=True)

# ===========================
# === İÇERİK ARŞİVİ MENÜSÜ ===
# ===========================

elif menu == "İçerik Arşivi":
    st.markdown("<h2 style='color:#8e54e9; font-weight:bold;'>İçerik Arşivi</h2>", unsafe_allow_html=True)
    st.info("Google Drive’daki medya, ürün görselleri ve kalite evraklarına aşağıdaki sekmelerden ulaşabilirsiniz.")

    # --- Klasör ID'leri (kolayca değiştirilebilir) ---
    DRIVE_FOLDER_IDS = {
        "Genel Medya Klasörü": "1gFAaK-6v1e3346e-W0TsizOqSq43vHLY",
        "Ürün Görselleri":      "18NNlmadm5NNFkI1Amzt_YMwB53j6AmbD",
        "Kalite Evrakları":     "1pbArzYfA4Tp50zvdyTzSPF2ThrMWrGJc",
    }

    def embed_url(folder_id: str) -> str:
        return f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"

    def open_url(folder_id: str) -> str:
        return f"https://drive.google.com/drive/folders/{folder_id}?usp=sharing"

    # --- Gömülü görünüm yüksekliği ayarı (ekranına göre) ---
    h = st.slider("Gömülü görünüm yüksekliği (px)", min_value=450, max_value=900, value=600, step=50)

    tabs = st.tabs(list(DRIVE_FOLDER_IDS.keys()))
    for tab, tab_name in zip(tabs, DRIVE_FOLDER_IDS.keys()):
        with tab:
            fid = DRIVE_FOLDER_IDS[tab_name]
            st.markdown(
                f"""
                <iframe src="{embed_url(fid)}"
                        width="100%" height="{h}" frameborder="0"
                        style="border:1px solid #eee; border-radius:12px; margin-top:10px;">
                </iframe>
                """,
                unsafe_allow_html=True
            )
            col_a, col_b = st.columns([1, 2])
            with col_a:
                st.link_button("Klasörü yeni sekmede aç", open_url(fid))
            with col_b:
                st.info("Dosya/klasörlere çift tıklayarak yeni sekmede açabilir veya indirebilirsiniz.")

    st.warning("Not: Klasörlerin paylaşımı 'Bağlantıya sahip olan herkes görüntüleyebilir' olmalı; aksi halde gömülü görünüm boş kalır.")


### ===========================
### --- SATIŞ ANALİTİĞİ MENÜSÜ ---
### ===========================

elif menu == "Satış Analitiği":
    st.markdown("<h2 style='color:#219A41; font-weight:bold;'>Satış Analitiği</h2>", unsafe_allow_html=True)

    # ---- Kolon güvenliği ----
    if "Tutar" not in df_evrak.columns:
        df_evrak["Tutar"] = 0
    date_col = "Fatura Tarihi" if "Fatura Tarihi" in df_evrak.columns else "Tarih"
    if date_col not in df_evrak.columns:
        df_evrak[date_col] = pd.NaT

    # ---- Tip dönüşümleri ----
    df_evrak = df_evrak.copy()
    df_evrak["Tutar_num"] = df_evrak["Tutar"].apply(smart_to_num).fillna(0.0)
    df_evrak[date_col] = pd.to_datetime(df_evrak[date_col], errors="coerce")
    df_evrak = df_evrak[df_evrak[date_col].notna()]  # geçersiz tarihleri at

    # ---- Toplamlar ----
    toplam_fatura = float(df_evrak["Tutar_num"].sum())
    st.markdown(
        f"<div style='font-size:1.3em; color:#185a9d; font-weight:bold;'>Toplam Fatura Tutarı: {toplam_fatura:,.2f} USD</div>",
        unsafe_allow_html=True,
    )
    # ---- Tarih aralığı filtresi (Timestamp ile) ----
    min_ts = df_evrak[date_col].min()
    max_ts = df_evrak[date_col].max()
    d1, d2 = st.date_input("Tarih Aralığı", value=(min_ts.date(), max_ts.date()))

    start_ts = pd.to_datetime(d1)  # 00:00
    end_ts   = pd.to_datetime(d2) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)  # gün sonu

    mask = df_evrak[date_col].between(start_ts, end_ts, inclusive="both")
    df_range = df_evrak[mask].copy()

    aralik_toplam = float(df_range["Tutar_num"].sum())
    st.markdown(
        f"<div style='font-size:1.2em; color:#f7971e; font-weight:bold;'>{d1} - {d2} Arası Toplam: {aralik_toplam:,.2f} USD</div>",
        unsafe_allow_html=True,
    )

    # ---- Müşteri filtresi ----
    df_analytics = df_range.copy()
    customer_col = "Müşteri Adı" if "Müşteri Adı" in df_analytics.columns else None

    selected_segment = None
    segment_label = "Müşteri Segmenti"
    if customer_col and "Kategori" in df_musteri.columns and not df_musteri.empty:
        segment_df = df_musteri[["Müşteri Adı", "Kategori"]].dropna(subset=["Müşteri Adı"]).copy()
        if not segment_df.empty:
            segment_df["Müşteri Adı"] = segment_df["Müşteri Adı"].astype(str).str.strip()
            segment_series = (
                segment_df.drop_duplicates("Müşteri Adı").set_index("Müşteri Adı")["Kategori"]
            )
            if not segment_series.empty:
                df_analytics[segment_label] = (
                    df_analytics[customer_col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .map(segment_series)
                    .fillna("Belirtilmemiş")
                )
                segment_options = (
                    ["Tüm Segmentler"]
                    + sorted(
                        df_analytics[segment_label]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist(),
                        key=lambda x: x.lower(),
                    )
                )
                if len(segment_options) > 1:
                    selected_segment = st.selectbox("Müşteri Segmenti", segment_options)
                    if selected_segment != "Tüm Segmentler":
                        df_analytics = df_analytics[df_analytics[segment_label] == selected_segment]

    df_filtered = df_analytics.copy()
    selected_customer = None
    if customer_col:
        musteri_listesi = sorted(
            df_analytics[customer_col]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", np.nan)
            .dropna()
            .unique()
        )
        musteri_opsiyonlari = ["Tüm Müşteriler"] + musteri_listesi
        selected_customer = st.selectbox("Müşteri Bazında Filtre", musteri_opsiyonlari)
        if selected_customer != "Tüm Müşteriler":
            df_filtered = df_filtered[
                df_filtered[customer_col].fillna("").astype(str).str.strip() == selected_customer
            ]

    filtered_total = float(df_filtered["Tutar_num"].sum())
    segment_text = (
        f"{selected_segment} Segmenti - "
        if selected_segment and selected_segment != "Tüm Segmentler"
        else ""
    )
    if customer_col and selected_customer and selected_customer != "Tüm Müşteriler":
        toplam_baslik = f"{segment_text}{selected_customer} Toplam"
    elif customer_col:
        toplam_baslik = f"{segment_text}Tüm Müşteriler Toplam"
    else:
        toplam_baslik = f"{segment_text}Seçili Aralık Toplam"

    st.markdown(
        f"<div style='font-size:1.1em; color:#185a9d; font-weight:bold;'>{toplam_baslik}: {filtered_total:,.2f} USD</div>",
        unsafe_allow_html=True,
    )
        # ---- En yüksek ciroya sahip müşteriler ----
    # ---- En yüksek ciroya sahip müşteriler ----
    if "Müşteri Adı" in df_analytics.columns and not df_analytics.empty:
        df_musteri = df_analytics.copy()
        df_musteri["Müşteri Adı"] = df_musteri["Müşteri Adı"].fillna("Bilinmeyen Müşteri")
        top_musteriler = (
            df_musteri.groupby("Müşteri Adı")["Tutar_num"]
            .sum()
            .sort_values(ascending=False)
            .head(5)
            .reset_index()
        )
        top_musteriler.rename(columns={"Tutar_num": "Toplam Ciro"}, inplace=True)

        st.markdown(
            "<h3 style='margin-top:20px; color:#185a9d;'>En Yüksek Ciroya Sahip İlk 5 Müşteri</h3>",
            unsafe_allow_html=True,
        )

        col_tab, col_chart = st.columns([1, 1])
        with col_tab:
            display_df = top_musteriler.copy()
            display_df["Toplam Ciro"] = display_df["Toplam Ciro"].map(lambda x: f"{x:,.2f} USD")
            st.dataframe(display_df, use_container_width=True)
        with col_chart:
            st.bar_chart(top_musteriler.set_index("Müşteri Adı")["Toplam Ciro"], use_container_width=True)
    else:
        st.info("Seçilen tarih aralığında müşteri bazlı ciro bilgisi bulunamadı.")

    if "Müşteri Adı" in df_analytics.columns and not df_analytics.empty:
        st.markdown(
            "<h3 style='margin-top:20px; color:#185a9d;'>Müşteri Bazında Ciro Yüzdeleri</h3>",
            unsafe_allow_html=True,
        )

        pie_df = df_analytics.copy()
        pie_df["Müşteri Adı"] = pie_df["Müşteri Adı"].fillna("Bilinmeyen Müşteri")
        pie_summary = (
            pie_df.groupby("Müşteri Adı")["Tutar_num"]
            .sum()
            .reset_index()
            .sort_values("Tutar_num", ascending=False)
        )

        if not pie_summary.empty:
            total_value = float(pie_summary["Tutar_num"].sum())


            if total_value <= 0:
                st.info("Müşteri bazında ciro yüzdesi hesaplanamadı.")
            else:
                pie_summary["Yüzde"] = pie_summary["Tutar_num"].apply(
                    lambda value: round((float(value) / total_value) * 100, 1) if total_value else 0.0
                )

                colors = plt.cm.tab20(np.linspace(0, 1, len(pie_summary)))
                fig, ax = plt.subplots(figsize=(8, 6))
                wedges, _, autotexts = ax.pie(
                    pie_summary["Tutar_num"],
                    autopct=lambda pct: f"%{pct:.1f}" if pct > 0 else "",
                    startangle=0,
                    colors=colors,
                    textprops={"color": "white", "weight": "bold"},
                )
                for autotext in autotexts:
                    autotext.set_fontsize(10)

                legend_labels = [
                    f"{label} (%{pct:.1f})" for label, pct in zip(pie_summary["Müşteri Adı"], pie_summary["Yüzde"])
                ]

                ax.legend(
                    wedges,
                    legend_labels,
                    title="Müşteriler",
                    loc="center left",
                    bbox_to_anchor=(1, 0.5),
                    fontsize=10,
                    title_fontsize=11,
                )
                ax.set_title("Müşteri Bazında Ciro Dağılımı", color="#185a9d", fontsize=14)
                ax.axis("equal")

                st.pyplot(fig, use_container_width=True)
                plt.close(fig)

                display_pie = pie_summary.copy()
                display_pie["Tutar (USD)"] = display_pie["Tutar_num"].map(lambda x: f"{float(x):,.2f}")
                display_pie["Yüzde (%)"] = display_pie["Yüzde"].map(lambda x: f"%{x:.1f}")
                display_pie = display_pie[["Müşteri Adı", "Tutar (USD)", "Yüzde (%)"]]
                st.dataframe(display_pie, use_container_width=True)
        else:
            st.info("Müşteri bazında ciro yüzdesi hesaplanamadı.")

    # ---- Detay tablo ----
    detail_cols = ["Müşteri Adı", "Fatura No", date_col, "Tutar"]
    detail_cols = [c for c in detail_cols if c in df_filtered.columns]

    if df_filtered.empty:
        st.info("Seçilen kriterlere uygun satış kaydı bulunamadı.")
    else:
        detail_df = df_filtered.copy()
        if date_col in detail_df.columns:
            detail_df = detail_df.sort_values(by=date_col, ascending=False)
        if detail_cols:
            detail_df = detail_df[detail_cols]
        st.dataframe(detail_df, use_container_width=True)

elif menu == "Help & Support":
    st.markdown(
        "<h2 style='color:#38bdf8; font-weight:bold;'>Yardım & Destek</h2>",
        unsafe_allow_html=True,
    )
    st.info(
        "Destek ekibimize ulaşmak için crm@sekeroglu.com.tr adresine e-posta gönderebilir veya +90 850 000 00 00 numaralı hattan iletişime geçebilirsiniz."
    )
    st.markdown(
        "<p style='color:#94a3b8;'>SSS dökümanımızı ve eğitim videolarımızı şirket içi paylaşım klasöründen inceleyebilirsiniz.</p>",
        unsafe_allow_html=True,
    )

elif menu == "Settings":
    st.markdown("<h2 style='color:#38bdf8; font-weight:bold;'>Ayarlar</h2>", unsafe_allow_html=True)
    st.info("Profil ayarları ve tema seçenekleri üzerinde çalışıyoruz. Güncellemeler yakında burada olacak.")

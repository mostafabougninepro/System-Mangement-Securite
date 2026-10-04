import io
import json
import os
import zipfile
import urllib.request
import csv
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pandas as pd
from PIL import Image as PILImage
import streamlit as st

# ================= ================= =================
# CONFIGURATION DE LA PAGE
# ================= ================= =================
st.set_page_config(
    page_title="ONCF — Management Sécurité",
    page_icon="🚄",
    layout="wide",
    initial_sidebar_state="expanded"
)

CLEAN_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
    .stApp { background-color: #F4F7F9; }
    .exec-header {
        background: linear-gradient(135deg, #0B1E36 0%, #16325B 100%);
        border-bottom: 4px solid #FF6B00;
        padding: 24px 32px;
        border-radius: 12px;
        box-shadow: 0 10px 20px rgba(11, 30, 54, 0.15);
        margin-bottom: 30px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .exec-title-wrapper { display: flex; align-items: baseline; gap: 15px; padding-left: 10px; }
    .exec-title-main { color: #FFFFFF; font-size: 22px; font-weight: 800; margin: 0; display: flex; align-items: center; gap: 10px; }
    .exec-title-sub { color: #FF8533; font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; border-left: 2px solid rgba(255, 255, 255, 0.3); padding-left: 12px; }
    .exec-badge { background: rgba(255, 107, 0, 0.2); color: #FF8533; border: 1px solid rgba(255, 107, 0, 0.4); padding: 6px 16px; border-radius: 30px; font-size: 12px; font-weight: 700; text-transform: uppercase; white-space: nowrap; }
    .stTextInput input, .stSelectbox select { border-radius: 6px !important; border: 1px solid #CBD5E1 !important; padding: 10px 14px !important; background-color: #FFFFFF !important; }
    .stButton>button { background: #0B1E36 !important; color: #FFFFFF !important; border-radius: 8px !important; border: none !important; font-weight: 700 !important; padding: 10px 20px !important; transition: all 0.2s !important; }
    .stButton>button:hover { background: #FF6B00 !important; transform: translateY(-2px) !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
</style>
"""

st.markdown(CLEAN_CSS, unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users_db.json")
PHOTOS_ZIP = os.path.join(BASE_DIR, "photos all.zip")
EXTRACTED_PHOTOS_DIR = os.path.join(BASE_DIR, "_extracted_photos")

# ================= LINKS GOOGLE SHEETS =================
GOOGLE_SHEET_CONDUITE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQz_31fa5jVtp2XvGWj3-HdKubKsG3AH2po3XSFJ8ZdKQs1f0nHRcgHG9Ah4zaziQ/pub?output=csv"
GOOGLE_SHEET_FORMATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQz_31fa5jVtp2XvGWj3-HdKubKsG3AH2po3XSFJ8ZdKQs1f0nHRcgHG9Ah4zaziQ/pub?output=csv&gid=1098236842"

# ================= ================= =================
# 1. BASE DE DONNEES UTILISATEURS
# ================= ================= =================
def load_users():
    if not os.path.exists(USERS_FILE):
        default_users = {"ADMIN": {"password": "adminpassword123", "role": "Admin", "nom": "Administrateur ONCF"}}
        save_users(default_users)
        return default_users
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=4)

st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("current_user", None)
st.session_state.setdefault("user_role", None)

# ================= ================= =================
# 2. LOGIN
# ================= ================= =================
if not st.session_state["logged_in"]:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("""
            <div style="background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); text-align: center; border-top: 4px solid #FF6B00;">
                <div style="font-size: 48px; margin-bottom: 10px;">🚄⚡</div>
                <h2 style="color: #0B1E36; font-weight: 800; margin: 0; font-size: 24px;">ONCF</h2>
                <p style="color: #64748B; font-size: 13px; font-weight: 600; text-transform: uppercase;">System management securite CCF.TC.Kenitra</p>
                <hr style="border: 0; height: 1px; background: #E2E8F0; margin: 20px 0;">
            </div>
        """, unsafe_allow_html=True)
        with st.form("login_form"):
            input_matricule = st.text_input("Matricule / Identifiant").strip().upper()
            input_password = st.text_input("Mot de passe", type="password")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("Se Connecter", use_container_width=True):
                users = load_users()
                if input_matricule in users and users[input_matricule]["password"] == input_password:
                    st.session_state["logged_in"] = True
                    st.session_state["current_user"] = input_matricule
                    st.session_state["user_role"] = users[input_matricule].get("role", "Utilisateur")
                    st.rerun()
                else:
                    st.error("❌ Identifiants incorrects.")
    st.stop()

# ================= ================= =================
# 3. HEADER & MENU
# ================= ================= =================
st.markdown(f"""
    <div class="exec-header">
        <div class="exec-title-wrapper">
            <div class="exec-title-main">🚄⚡ Office National des Chemins de Fer</div>
            <div class="exec-title-sub">System management securite CCF.TC.Kenitra</div>
        </div>
        <div style="display: flex; align-items: center; gap: 15px;">
            <div class="exec-badge">{st.session_state['user_role']}</div>
            <div style="color: white; font-size: 14px; font-weight: 600;">👤 {st.session_state['current_user']}</div>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("### 🏛 Navigation")
if st.session_state["user_role"] == "Admin":
    menu = st.sidebar.radio("Module actif :", ["🪪 Cartes d'Habilitation", "👥 Gestion des Accès"])
else:
    menu = "🪪 Cartes d'Habilitation"

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Déconnexion", use_container_width=True):
    st.session_state["logged_in"] = False
    st.session_state["current_user"] = None
    st.session_state["user_role"] = None
    st.rerun()

# ================= ================= =================
# 4. GESTION DES ACCES (ADMIN)
# ================= ================= =================
if menu == "👥 Gestion des Accès":
    st.markdown("### 👥 Administration des Utilisateurs")
    col_u1, col_u2 = st.columns([1, 1.2])
    with col_u1:
        st.markdown("##### ➕ Nouvel Utilisateur")
        users = load_users()
        with st.form("add_user_form"):
            new_mat = st.text_input("Matricule").strip().upper()
            new_pass = st.text_input("Mot de passe", type="password")
            new_role = st.selectbox("Rôle Système", ["Utilisateur", "Admin"])
            if st.form_submit_button("Créer l'utilisateur", use_container_width=True):
                if not new_mat or not new_pass:
                    st.error("Champs obligatoires non renseignés.")
                elif new_mat in users:
                    st.warning("Utilisateur déjà existant.")
                else:
                    users[new_mat] = {"password": new_pass, "role": new_role}
                    save_users(users)
                    st.success(f"Compte {new_mat} créé avec succès !")
                    st.rerun()
    with col_u2:
        st.markdown("##### 📋 Comptes Enregistrés")
        users_list = [{"Matricule": m, "Rôle": d.get("role", "Utilisateur")} for m, d in users.items()]
        st.dataframe(pd.DataFrame(users_list), use_container_width=True)
    st.stop()

# ================= ================= =================
# 5. GENERATEUR DE CARTES D'HABILITATION
# ================= ================= =================
def get_agent_photo(matricule):
    if not matricule or not str(matricule).strip():
        return None, "Matricule vide"
    if os.path.exists(PHOTOS_ZIP):
        if not os.path.exists(EXTRACTED_PHOTOS_DIR):
            os.makedirs(EXTRACTED_PHOTOS_DIR, exist_ok=True)
            try:
                with zipfile.ZipFile(PHOTOS_ZIP, 'r') as zip_ref:
                    zip_ref.extractall(EXTRACTED_PHOTOS_DIR)
            except Exception:
                pass
    target = str(matricule).strip().lower()
    if os.path.exists(EXTRACTED_PHOTOS_DIR):
        try:
            for root, _, files in os.walk(EXTRACTED_PHOTOS_DIR):
                for file_name in files:
                    name_part, _ = os.path.splitext(file_name)
                    if name_part.strip().lower() == target:
                        return os.path.join(root, file_name), "Photo trouvée"
        except Exception:
            pass
    return None, "Photo non trouvable"

def get_agent_dates_and_details(matricule, page_type="conduite"):
    try:
        target_url = GOOGLE_SHEET_FORMATION_URL if page_type.lower() == "formation" else GOOGLE_SHEET_CONDUITE_URL
        
        req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            content = response.read().decode('utf-8')
            
        reader = csv.reader(io.StringIO(content))
        for row in reader:
            if not row or len(row) <= 2:
                continue
            
            row_matricule = str(row[2]).strip() if len(row) > 2 else ""
            if row_matricule.lower() == str(matricule).strip().lower():
                
                def fmt_date(val):
                    if val and str(val).lower() not in ["nan", "nat", "", "néant", "None"]:
                        parsed_date = pd.to_datetime(val, errors='coerce', dayfirst=True)
                        if pd.notnull(parsed_date):
                            return parsed_date.strftime("%Y-%m-%d")
                        return str(val).strip()
                    return ""

                nom_val = str(row[0]).strip() if len(row) > 0 else ""      # A (Nom)
                prenom_val = str(row[1]).strip() if len(row) > 1 else ""   # B (Prénom)
                mat_val = str(row[2]).strip() if len(row) > 2 else ""      # C (Matricule)
                dt_auth = fmt_date(row[4]) if len(row) > 4 else ""         # E
                fonction_val = str(row[12]).strip() if len(row) > 12 else "" # M (Fonction)
                engin_val = str(row[13]).strip() if len(row) > 13 else ""  # N
                ligne_site_val = str(row[14]).strip() if len(row) > 14 else "" # O
                dt_med = fmt_date(row[19]) if len(row) > 19 else ""        # T
                dt_psy = fmt_date(row[24]) if len(row) > 24 else ""        # Y
                dt_prof = fmt_date(row[27]) if len(row) > 27 else ""       # AB

                return {
                    "Nom": nom_val,
                    "Prenom": prenom_val,
                    "Matricule": mat_val,
                    "Fonction": fonction_val,
                    "Date_Autorisation": dt_auth,
                    "Examen_Medical": dt_med,
                    "Examen_Psychotechnique": dt_psy,
                    "Examen_Professionnel": dt_prof,
                    "Engin": engin_val,
                    "Ligne_Site": ligne_site_val,
                }
    except Exception as e:
        st.error(f"Erreur: {e}")
    return {}
    
def determine_template_and_mapping(fonction):
    f_lower = fonction.lower().strip()
    
    if "conducteur de ligne" in f_lower:
        keyword = "cl"
        default_eng = "E1450 , E1400 , Z2M"
        default_sit = ""
    elif "chef de trains" in f_lower or "chef de train" in f_lower:
        keyword = "ctr"
        default_eng = "E1450 , E1400 , E1250 , Z2M , DH400"
        default_sit = ""
    elif "conducteur de manœuvre" in f_lower or "conducteur de manoeuvre" in f_lower:
        keyword = "crmv"
        default_eng = "E1450 , E1400 , Z2M , DH400 , DH350 , DM600"
        default_sit = " Site Voyageurs Kénitra "
    elif "chef formation trains" in f_lower:
        keyword = "cft"
        default_eng = "E1450 , E1400 , E1250 , Z2M , DH400 , DM600"
        default_sit = " Site Voyageurs Kénitra "
    else:
        keyword = "ctr"
        default_eng = "E1450 , E1400 , E1250 , Z2M , DH400"
        default_sit = ""

    matched_file = keyword + ".xlsx"
    for f in os.listdir(BASE_DIR):
        if f.lower().endswith(".xlsx") and keyword in f.lower():
            matched_file = f
            break

    return {
        "keyword": keyword.upper(),
        "template": matched_file,
        "default_engins": default_eng,
        "default_site": default_sit,
    }

st.markdown("### 🔍 Recherche & Identification de l'Agent")

st.session_state.setdefault("last_matricule", "")
st.session_state.setdefault("last_page_type", "Conduite")

selected_page = st.selectbox("Type de Registre", ["Conduite", "Formation"])
matricule_search = st.text_input("Saisir le Matricule de l'agent :", placeholder="Exemple: 42685P")

if matricule_search != st.session_state["last_matricule"] or selected_page != st.session_state["last_page_type"]:
    st.session_state["last_matricule"] = matricule_search
    st.session_state["last_page_type"] = selected_page
    
    dates_info = get_agent_dates_and_details(matricule_search, page_type=selected_page) if matricule_search else {}

    st.session_state["matricule"] = dates_info.get("Matricule", matricule_search)
    st.session_state["nom"] = dates_info.get("Nom", "")
    st.session_state["prenom"] = dates_info.get("Prenom", "")
    st.session_state["fonction"] = dates_info.get("Fonction", "Chef de Train" if selected_page == "Conduite" else "Chef Formation Trains")

    config_info = determine_template_and_mapping(st.session_state["fonction"])
    st.session_state["dt_auth"] = dates_info.get("Date_Autorisation", "")
    st.session_state["dt_med"] = dates_info.get("Examen_Medical", "")
    st.session_state["dt_psy"] = dates_info.get("Examen_Psychotechnique", "")
    st.session_state["dt_prof"] = dates_info.get("Examen_Professionnel", "")
    st.session_state["lignes"] = dates_info.get("Ligne_Site") or config_info["default_site"]
    st.session_state["engins"] = dates_info.get("Engin") or config_info["default_engins"]

found_photo_path, search_status = get_agent_photo(matricule_search)
final_photo_source = None

col_p1, col_p2 = st.columns([1, 3])
with col_p1:
    uploaded_photo = st.file_uploader("Photo d'identité (Optionnel)", type=["jpg", "jpeg", "png"])
    if uploaded_photo is not None:
        final_photo_source = uploaded_photo
        st.image(uploaded_photo, caption="Photo importée", width=115)
    elif found_photo_path:
        final_photo_source = found_photo_path
        st.image(found_photo_path, caption=f"✅ {search_status}", width=115)
    elif matricule_search.strip():
        st.warning("⚠️ Photo non disponible")

st.markdown("---")
st.markdown("### 📝 Informations d'Habilitation")

col1, col2 = st.columns(2)
with col1:
    nom_input = st.text_input("Nom", key="nom")
    matricule_input = st.text_input("Matricule", key="matricule")
    fonction_input = st.text_input("Fonction (Titre d'habilitation)", key="fonction")
    dt_autorisation = st.text_input("Date d'autorisation", key="dt_auth")
    dt_medical = st.text_input("Date d'expiration de l'examen médical", key="dt_med")

with col2:
    prenom_input = st.text_input("Prénom", key="prenom")
    dt_professionnel = st.text_input("Date d'expiration de l'examen pro", key="dt_prof")
    dt_psycho = st.text_input("Date d'expiration de l'examen psycho", key="dt_psy")

lignes_sites = st.text_input("Autorisé aux lignes suivantes", key="lignes")
materiel_locos = st.text_input("Autorisé à arrêter أو Autorisé à conduire", key="engins")

st.markdown("<br>", unsafe_allow_html=True)

def generate_custom_excel():
    config = determine_template_and_mapping(fonction_input)
    tmpl_filename = config["template"]
    
    tmpl_path = os.path.join(BASE_DIR, tmpl_filename)
    if not os.path.exists(tmpl_path):
        for f in os.listdir(BASE_DIR):
            if f.endswith(".xlsx") and any(k in f.lower() for k in ["ctr", "cl", "cft", "crmv"]):
                tmpl_path = os.path.join(BASE_DIR, f)
                break

    wb = openpyxl.load_workbook(tmpl_path)
    sheet = wb.active

    sheet["E6"] = nom_input
    sheet["J6"] = prenom_input
    sheet["E7"] = matricule_input
    sheet["G10"] = dt_autorisation
    sheet["G11"] = dt_professionnel
    sheet["G12"] = dt_medical
    sheet["G13"] = dt_psycho

    sheet["L5"] = materiel_locos
    sheet["L5"].alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center", horizontal="center")

    sheet["M5"] = lignes_sites
    sheet["M5"].alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center", horizontal="center")

    if final_photo_source is not None:
        pil_img = PILImage.open(final_photo_source if isinstance(final_photo_source, str) else io.BytesIO(final_photo_source.read()))
        
        # ضبط القياسات بدقة باش متخرجش على الخلايا C6, C7, C8
        target_w, target_h = int(1.8 * 37.8), int(1.3 * 37.8)
        pil_img = pil_img.resize((target_w, target_h), PILImage.Resampling.LANCZOS)
        
        img_temp_path = os.path.join(BASE_DIR, "_temp_photo.png")
        pil_img.save(img_temp_path)

        xl_img = OpenpyxlImage(img_temp_path)
        xl_img.width = target_w
        xl_img.height = target_h
        sheet.add_image(xl_img, "B5")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if st.button("⚡ Générer la Carte d'Habilitation", use_container_width=True):
    excel_file = generate_custom_excel()
    config = determine_template_and_mapping(fonction_input)
    card_type_prefix = config["keyword"]
    clean_nom = " ".join(nom_input.strip().upper().split()) if nom_input.strip() else "AGENT"
    file_download_name = f"Carte_{card_type_prefix}_{clean_nom}.xlsx"

    st.success(f"✅ Document d'habilitation prêt : {file_download_name}")
    st.download_button(
        label=f"📥 Télécharger {file_download_name}",
        data=excel_file,
        file_name=file_download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

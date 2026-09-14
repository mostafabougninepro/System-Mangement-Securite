import io
import json
import os
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pandas as pd
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users_db.json")
PHOTOS_DIR = os.path.join(BASE_DIR, "photos")

# ================= ================= =================
# CSS PROPRE & DESIGN ONCF
# ================= ================= =================
CLEAN_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
    }

    .stApp {
        background-color: #F4F7F9;
    }

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

    .exec-title-wrapper {
        display: flex;
        align-items: baseline;
        gap: 15px;
        padding-left: 10px;
    }

    .exec-title-main {
        color: #FFFFFF;
        font-size: 22px;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .exec-title-sub {
        color: #FF8533;
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        border-left: 2px solid rgba(255, 255, 255, 0.3);
        padding-left: 12px;
    }

    .exec-badge {
        background: rgba(255, 107, 0, 0.2);
        color: #FF8533;
        border: 1px solid rgba(255, 107, 0, 0.4);
        padding: 6px 16px;
        border-radius: 30px;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        white-space: nowrap;
    }

    .stTextInput input, .stSelectbox select {
        border-radius: 6px !important;
        border: 1px solid #CBD5E1 !important;
        padding: 10px 14px !important;
        background-color: #FFFFFF !important;
    }

    .stButton>button {
        background: #0B1E36 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
        transition: all 0.2s !important;
    }

    .stButton>button:hover {
        background: #FF6B00 !important;
        transform: translateY(-2px) !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
"""
st.markdown(CLEAN_CSS, unsafe_allow_html=True)

# ================= ================= =================
# 1. BASE DE DONNEES UTILISATEURS
# ================= ================= =================
def load_users():
    if not os.path.exists(USERS_FILE):
        default_users = {
            "ADMIN": {
                "password": "adminpassword123",
                "role": "Admin",
                "nom": "Administrateur ONCF"
            }
        }
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
# 2. MODULE DE CONNEXION
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
            submit_login = st.form_submit_button("Se Connecter", use_container_width=True)
            
            if submit_login:
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
            <div style="color: white; font-size: 14px; font-weight: 600;">
                👤 {st.session_state['current_user']}
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("### 🏛️ Navigation")
if st.session_state["user_role"] == "Admin":
    menu = st.sidebar.radio("Module actif :", ["🪪 Cartes d'Habilitation", "👥 Gestion des Accès"])
else:
    menu = "🪪 Cartes d'Habilitation"

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Session")
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
    
    target = str(matricule).strip().lower()
    # تنظيف الماتريكول من الرموز للمقارنة المرنة
    clean_target = "".join(c for c in target if c.isalnum()).lower()
    
    if os.path.exists(PHOTOS_DIR):
        try:
            for root, dirs, files in os.walk(PHOTOS_DIR):
                for file_name in files:
                    name_part, ext = os.path.splitext(file_name)
                    if ext.lower() not in ['.jpg', '.jpeg', '.png']:
                        continue
                    clean_name_part = "".join(c for c in name_part if c.isalnum()).lower()
                    
                    # مطابقة دقيقة أو جزئية
                    if clean_target == clean_name_part or clean_target in clean_name_part or clean_name_part in clean_target:
                        return os.path.join(root, file_name), "Photo trouvée"
        except Exception:
            pass
    return None, "Photo non trouvable"

def get_agent_info_complet(matricule):
    excel_filenames = [
        f for f in os.listdir(BASE_DIR) 
        if f.lower().endswith(".xlsx") 
        and f.lower() not in ["cft.xlsx", "cl.xlsx", "crmv.xlsx", "ctr.xlsx"]
    ]
    
    info = {
        "Nom": "", "Prenom": "", "Fonction": "Chef de Trains",
        "Date_Autorisation": "", "Examen_Medical": "",
        "Examen_Psychotechnique": "", "Examen_Professionnel": "",
        "Engin": "", "Ligne_Site": ""
    }
    
    if not excel_filenames or not str(matricule).strip():
        return info
    
    target = str(matricule).strip().lower()
    
    for excel_filename in excel_filenames:
        excel_path = os.path.join(BASE_DIR, excel_filename)
        try:
            xl = pd.ExcelFile(excel_path)
            for sheet_name in xl.sheet_names:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                
                mle_col = next((c for c in df.columns if str(c).strip().lower() in ["matricule", "mle", "mat", "N° matricule"]), None)
                if not mle_col:
                    try:
                        df_header6 = pd.read_excel(excel_path, sheet_name=sheet_name, header=6)
                        df_header6.columns = [str(c).strip() for c in df_header6.columns]
                        mle_col = next((c for c in df_header6.columns if str(c).strip().lower() in ["matricule", "mle", "mat"]), None)
                        if mle_col:
                            df = df_header6
                    except Exception:
                        pass
                
                if mle_col:
                    df[mle_col] = df[mle_col].astype(str).str.strip()
                    agent = df[df[mle_col].str.lower() == target]
                    
                    if not agent.empty:
                        data = agent.iloc[0]
                        def fmt_date(val):
                            if pd.notnull(val) and str(val) != "NaT" and str(val).strip() != "":
                                dt_parsed = pd.to_datetime(val, errors='coerce')
                                if pd.notnull(dt_parsed):
                                    return dt_parsed.strftime("%Y-%m-%d")
                                return str(val).strip()
                            return ""

                        info["Nom"] = next((str(data[c]).strip() for c in df.columns if str(c).lower().strip() in ["nom", "nom & prénom", "nom et prénom"] and pd.notnull(data[c])), info["Nom"])
                        info["Prenom"] = next((str(data[c]).strip() for c in df.columns if str(c).lower().strip() in ["prénom", "prenom"] and pd.notnull(data[c])), info["Prenom"])
                        
                        fct = next((str(data[c]).strip() for c in df.columns if any(k in str(c).lower() for k in ["fonction", "titre", "emploi"]) and pd.notnull(data[c])), "")
                        if fct and fct.lower() != "nan": info["Fonction"] = fct
                        
                        info["Ligne_Site"] = next((str(data[c]).strip() for c in df.columns if any(k in str(c).lower() for k in ["ligne", "site", "parcours"]) and pd.notnull(data[c]) and str(data[c]).lower() != "nan"), "")
                        info["Engin"] = next((str(data[c]).strip() for c in df.columns if any(k in str(c).lower() for k in ["engin", "materiel", "loco", "rame"]) and pd.notnull(data[c]) and str(data[c]).lower() != "nan"), "")
                        
                        for col in df.columns:
                            col_l = str(col).lower()
                            val_cell = data[col]
                            if pd.isnull(val_cell) or str(val_cell).strip() == "" or str(val_cell).lower() == "nan":
                                continue
                            
                            formatted_d = fmt_date(val_cell)
                            if formatted_d:
                                if any(k in col_l for k in ["autorisation", "delivrance", "date d'autorisation", "emis"]):
                                    if not info["Date_Autorisation"]: info["Date_Autorisation"] = formatted_d
                                elif any(k in col_l for k in ["médical", "medical", "visite", "vm"]):
                                    if not info["Examen_Medical"]: info["Examen_Medical"] = formatted_d
                                elif any(k in col_l for k in ["psy", "psychotechnique"]):
                                    if not info["Examen_Psychotechnique"]: info["Examen_Psychotechnique"] = formatted_d
                                elif any(k in col_l for k in ["professionnel", "prof", "evaluation", "eval"]):
                                    if not info["Examen_Professionnel"]: info["Examen_Professionnel"] = formatted_d

                        return info
        except Exception:
            continue
    return info

def determine_template_and_defaults(fonction):
    f_lower = str(fonction).lower().strip()
    if "manœuvre" in f_lower or "manoeuvre" in f_lower or "crmv" in f_lower:
        return "CRMV.xlsx", "E1450, E1400, Z2M, DH400, DH350, DM600", "Site Voyageurs Kénitra"
    elif "formation" in f_lower or "cft" in f_lower:
        return "CFT.xlsx", "E1450, E1400, E1250, Z2M, DH400, DM600", "Site Voyageurs Kénitra"
    elif "ligne" in f_lower or "cl" in f_lower:
        return "CL.xlsx", "E1450, E1400, Z2M", ""
    else:
        return "CTR.xlsx", "E1450, E1400, E1250, Z2M, DH400", ""

# SECTION: RECHERCHE
st.markdown("### 🔍 Recherche & Identification de l'Agent")
st.session_state.setdefault("last_matricule", "")
matricule_search = st.text_input("Saisir le Matricule de l'agent :", placeholder="Exemple: 47622S")

if matricule_search != st.session_state["last_matricule"]:
    st.session_state["last_matricule"] = matricule_search
    agent_info = get_agent_info_complet(matricule_search)

    st.session_state["nom"] = agent_info["Nom"]
    st.session_state["prenom"] = agent_info["Prenom"]
    st.session_state["matricule"] = matricule_search
    st.session_state["fonction"] = agent_info["Fonction"]

    _, def_engins, def_site = determine_template_and_defaults(st.session_state["fonction"])
    
    st.session_state["dt_auth"] = agent_info["Date_Autorisation"]
    st.session_state["dt_med"] = agent_info["Examen_Medical"]
    st.session_state["dt_psy"] = agent_info["Examen_Psychotechnique"]
    st.session_state["dt_prof"] = agent_info["Examen_Professionnel"]
    st.session_state["lignes"] = agent_info["Ligne_Site"] or def_site
    st.session_state["engins"] = agent_info["Engin"] or def_engins

found_photo_path, search_status = get_agent_photo(matricule_search)

col_p1, col_p2 = st.columns([1, 2])
with col_p1:
    uploaded_file = st.file_uploader("Photo d'identité (Optionnel)", type=["jpg", "jpeg", "png"])

active_photo_path = None
if uploaded_file is not None:
    temp_uploaded_path = os.path.join(BASE_DIR, "_uploaded_temp_photo.png")
    with open(temp_uploaded_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    active_photo_path = temp_uploaded_path
elif found_photo_path:
    active_photo_path = found_photo_path

with col_p2:
    if active_photo_path:
        st.image(active_photo_path, width=115, caption="✅ Photo prête")
    else:
        st.warning("⚠️ Photo non trouvée. Vous pouvez l'importer manuellement.")

st.markdown("---")

# SECTION: INFORMATIONS
st.markdown("### 📝 Informations d'Habilitation")

col1, col2 = st.columns(2)
with col1:
    nom_input = st.text_input("Nom", key="nom")
    matricule_input = st.text_input("Matricule", key="matricule")
    fonction_input = st.text_input("Fonction (Titre d'habilitation)", key="fonction")
    dt_autorisation = st.text_input("Date d'autorisation", key="dt_auth")
    dt_medical = st.text_input("Date examen médical", key="dt_med")

with col2:
    prenom_input = st.text_input("Prénom", key="prenom")
    dt_professionnel = st.text_input("Date examen professionnel", key="dt_prof")
    dt_psycho = st.text_input("Date examen psychotechnique", key="dt_psy")

lignes_sites = st.text_input("Lignes / Sites autorisés", key="lignes")
materiel_locos = st.text_input("Matériel / Locos / Rames", key="engins")

st.markdown("<br>", unsafe_allow_html=True)

def generate_custom_excel():
    tmpl_filename, _, _ = determine_template_and_defaults(fonction_input)
    tmpl_path = next((os.path.join(BASE_DIR, f) for f in os.listdir(BASE_DIR) if f.lower() == tmpl_filename.lower()), os.path.join(BASE_DIR, tmpl_filename))

    wb = openpyxl.load_workbook(tmpl_path)
    sheet = wb.active

    sheet["D4"] = fonction_input
    sheet["F5"] = nom_input
    sheet["J5"] = prenom_input
    sheet["F6"] = matricule_input
    sheet["F8"] = dt_autorisation
    sheet["F9"] = dt_professionnel
    sheet["F10"] = dt_medical
    sheet["F11"] = dt_psycho
    sheet["L4"] = materiel_locos
    sheet["Q4"] = lignes_sites

    if active_photo_path and os.path.exists(active_photo_path):
        try:
            xl_img = OpenpyxlImage(active_photo_path)
            sheet.add_image(xl_img, "B5")
        except Exception:
            pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if st.button("⚡ Générer la Carte d'Habilitation", use_container_width=True):
    excel_file = generate_custom_excel()
    clean_nom = nom_input.strip().upper() if nom_input.strip() else "AGENT"
    file_download_name = f"Carte_{clean_nom}.xlsx"

    st.success(f"✅ Document d'habilitation prêt : {file_download_name}")
    st.download_button(
        label=f"📥 Télécharger {file_download_name}",
        data=excel_file,
        file_name=file_download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

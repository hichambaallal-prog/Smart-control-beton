import datetime
514f8421-af39-425a-87e4-25cc7b640832.py
Python
est ce que  quand j'ai scanné un QR  de ne pas demander chaque fois le mot de passe sachant que je suis déjà connecter a la plateforme depuis mon smartphone .chaque fois que j'ai scanné QR demande mot de passe .


Oui, c’est tout à fait possible, et dans ton cas je vois exactement pourquoi le mot de passe est redemandé.

Pourquoi cela arrive actuellement

Dans ton code, la connexion est conservée uniquement dans st.session_state :

if "user" not in st.session_state:
    st.session_state["user"] = None

Puis l'accès est autorisé seulement si :

if st.session_state["user"] is None:
    # afficher la page de connexion

Donc, lorsque tu scannes un QR Code, le navigateur ouvre une nouvelle session Streamlit. La plateforme ne retrouve plus st.session_state["user"], même si tu étais déjà connecté auparavant.

Ton code QR fonctionne bien : il récupère rec, beton_id et ep, mais ces informations sont également placées dans session_state.

Ce que je te recommande

Pour ta plateforme LPEE, je ferais le fonctionnement suivant :

1. Première connexion

Nom utilisateur + mot de passe → connexion.

2. La plateforme mémorise la connexion sur le smartphone

création d'une session persistante sécurisée.

3. Tu scannes un QR Code

QR → ouverture de la fiche directement.

4. Aucun mot de passe demandé

puisque le smartphone possède déjà une session valide.

Par exemple :

Connexion initiale

🔐 Accès Restreint - LPEE
BAALLAL + mot de passe
→ Se connecter

Puis :

Scan QR

📱 Scan QR Code
→ vérification automatique de la session
→ Suivi Contrôle Béton
→ ouverture directement de la fiche concernée.

⚠️ Important

Il ne faut pas mettre le mot de passe dans le QR Code. Le QR doit uniquement contenir l'identifiant de la fiche, comme tu le fais déjà avec :

?rec=...&beton_id=...

La bonne solution est donc de remplacer la connexion uniquement basée sur st.session_state par une authentification persistante par cookie/session sécurisée.

Et il y a un autre point important dans ton fichier : les mots de passe des utilisateurs apparaissent actuellement directement dans le code source (DEFAULT_USERS). Je te conseille fortement de les retirer du code et de les changer, surtout si ce projet est déployé en ligne.

Si tu veux, je peux modifier directement ton fichier Python pour que :

Connexion une seule fois → scan QR → accès direct à la fiche sans redemander le mot de passe, tout en conservant les rôles admin, laboratoire, bétonnage, etc.

envoyer moi le fichier python avec cette mise ajour .

Oui. J’ai intégré la session persistante dans ton fichier Python.

Fonctionnement maintenant
Première connexion : nom d’utilisateur + mot de passe.
La session est mémorisée pendant 30 jours sur le navigateur du smartphone.
Tu scans ensuite un QR Code → pas de nouvelle demande de mot de passe.
Le QR Code continue à ouvrir directement la fiche concernée.
La session est liée au navigateur utilisé : si tu changes de navigateur, il faudra se reconnecter.
Le bouton Déconnexion supprime également la session mémorisée.
Le système conserve les rôles et droits existants de ta plateforme.
📥 Fichier Python mis à jour

Télécharger le fichier Python – LPEE avec connexion persistante

J’ai également vérifié que le fichier est syntaxiquement valide (py_compile OK).

Important : pour une utilisation en production, je recommande de définir une valeur AUTH_SECRET forte dans les secrets Streamlit. Le système utilise ce secret pour signer les sessions persistantes.

LPEE_CTR_CSB_plateforme_session_persistante.py
Code

Bibliothèque
/
LPEE_CTR_CSB_plateforme_session_persistante.py
import datetime
import os
import base64
import hashlib
import hmac
import json
import time
from fpdf import FPDF
from PIL import Image
import streamlit as st
import streamlit.components.v1 as components
from supabase import Client, create_client

# Importation sécurisée du gestionnaire Hors-Ligne SQLite
try:
  from offline_manager import (
      get_pending_count,
      init_offline_db,
      insert_safe,
      sync_data_to_supabase,
  )

  init_offline_db()
  OFFLINE_SUPPORT = True
except ImportError:
  OFFLINE_SUPPORT = False

# ==========================================
# 1. CONFIGURATION DE LA PAGE & INJECTION PWA
# ==========================================
st.set_page_config(
    page_title="LPEE - CTR-CSB",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================
# 1bis. AUTHENTIFICATION PERSISTANTE (30 jours)
# ==========================================
# Après une première connexion réussie, un jeton signé est conservé
# dans le navigateur. Ainsi, un scan QR dans le même smartphone peut
# rouvrir la plateforme sans redemander le mot de passe.
# Le mot de passe n'est JAMAIS stocké dans le QR Code.
try:
  AUTH_SECRET = st.secrets.get("AUTH_SECRET", "")
except Exception:
  AUTH_SECRET = ""

# En production, définir AUTH_SECRET dans les secrets Streamlit.
if not AUTH_SECRET:
  AUTH_SECRET = os.environ.get("AUTH_SECRET", "LPEE-CTR-CSB-CHANGE-ME")

AUTH_TOKEN_TTL = 30 * 24 * 60 * 60  # 30 jours
AUTH_STORAGE_KEY = "lpee_auth_token"

def _b64encode(data: bytes) -> str:
  return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")

def _b64decode(value: str) -> bytes:
  return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))

def create_auth_token(username: str) -> str:
  payload = {"u": username, "exp": int(time.time()) + AUTH_TOKEN_TTL}
  raw = json.dumps(payload, separators=(",", ":")).encode("utf-8")
  payload_b64 = _b64encode(raw)
  signature = hmac.new(
      AUTH_SECRET.encode("utf-8"), payload_b64.encode("utf-8"), hashlib.sha256
  ).digest()
  return payload_b64 + "." + _b64encode(signature)

def verify_auth_token(token: str):
  try:
    payload_b64, signature_b64 = token.split(".", 1)
    expected = hmac.new(
        AUTH_SECRET.encode("utf-8"), payload_b64.encode("utf-8"), hashlib.sha256
    ).digest()
    received = _b64decode(signature_b64)
    if not hmac.compare_digest(expected, received):
      return None
    payload = json.loads(_b64decode(payload_b64).decode("utf-8"))
    if int(payload.get("exp", 0)) < int(time.time()):
      return None
    username = str(payload.get("u", "")).strip().upper()
    return username or None
  except Exception:
    return None

# Si le navigateur nous renvoie un jeton mémorisé, on prépare
# l'ouverture automatique de la session. La validation de l'utilisateur
# est faite juste après le chargement de la base des utilisateurs.
_auth_from_url = st.query_params.get("auth_token")
if _auth_from_url:
  _auth_username = verify_auth_token(str(_auth_from_url))
  if _auth_username:
    st.session_state["pending_auth_username"] = _auth_username

# ==========================================
# 1ter. CAPTURE DES PARAMÈTRES QR CODE (ex: ?rec=...&beton_id=...)
# ==========================================
# C'est ICI qu'il faut lire l'URL, et pas dans views/suivi_controle_beton.py :
# ce module est importé (pas exécuté directement), donc son bloc
# "if __name__ == '__main__':" ne s'exécute jamais dans l'app déployée.
_query_params = st.query_params
_qr_rec = _query_params.get("rec") or _query_params.get("num_reception")
_qr_bid = _query_params.get("beton_id") or _query_params.get("id")
_qr_ep = _query_params.get("ep")

if _qr_rec or _qr_bid:
    if _qr_rec:
        st.session_state["pending_qr_rec"] = str(_qr_rec).strip()
    if _qr_bid:
        st.session_state["pending_qr_bid"] = str(_qr_bid).strip()
    if _qr_ep:
        st.session_state["pending_qr_ep"] = str(_qr_ep).strip()

    # (Ré)armer la redirection automatique vers la page "Suivi Contrôle Béton"
    st.session_state["qr_page_applied"] = False

    # Nettoyer uniquement les paramètres QR. Le jeton d'authentification
    # est géré séparément afin de ne pas perdre la session persistante.
    for _key in ["rec", "num_reception", "beton_id", "id", "ep"]:
      if _key in st.query_params:
        del st.query_params[_key]

# Injection PWA dans le HEAD du document principal
pwa_code = """
<script>
const parentDoc = window.parent.document;

if (!parentDoc.querySelector('link[rel="manifest"]')) {
    const manifestLink = parentDoc.createElement('link');
    manifestLink.rel = 'manifest';
    manifestLink.href = '/manifest.json';
    parentDoc.head.appendChild(manifestLink);
}

if (!parentDoc.querySelector('meta[name="theme-color"]')) {
    const metaTheme = parentDoc.createElement('meta');
    metaTheme.name = 'theme-color';
    metaTheme.content = '#0066cc';
    parentDoc.head.appendChild(metaTheme);
}

if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/sw.js')
        .then((reg) => console.log('Service Worker PWA enregistré !', reg))
        .catch((err) => console.error('Erreur Service Worker PWA :', err));
}

// Restauration automatique de la session après un scan QR.
// Le token est stocké uniquement dans localStorage du navigateur.
try {
    const STORAGE_KEY = "lpee_auth_token";
    const url = new URL(window.parent.location.href);
    if (url.searchParams.get("logout") === "1") {
        window.localStorage.removeItem(STORAGE_KEY);
        url.searchParams.delete("logout");
        window.parent.history.replaceState({}, document.title, url.toString());
    }
    const token = window.localStorage.getItem(STORAGE_KEY);
    const hasToken = url.searchParams.has("auth_token");
    if (token && !hasToken) {
        url.searchParams.set("auth_token", token);
        window.parent.location.replace(url.toString());
    }

    // Après connexion, le serveur place auth_token dans l'URL.
    // On le mémorise puis on nettoie l'URL.
    if (hasToken) {
        const urlToken = url.searchParams.get("auth_token");
        if (urlToken) {
            window.localStorage.setItem(STORAGE_KEY, urlToken);
            url.searchParams.delete("auth_token");
            window.parent.history.replaceState({}, document.title, url.toString());
        }
    }
} catch (e) {
    console.warn("Session persistante indisponible :", e);
}
</script>
"""
components.html(pwa_code, height=0, width=0)


# ==========================================
# 2. MODULE DE GÉNÉRATION PDF (FPDF2)
# ==========================================
class LPEEPDFReport(FPDF):

  def header(self):
    self.set_font("Helvetica", "B", 14)
    self.cell(
        0,
        8,
        "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES - LPEE",
        border=False,
        new_x="LMARGIN",
        new_y="NEXT",
        align="C",
    )
    self.set_font("Helvetica", "B", 11)
    self.cell(
        0,
        6,
        "CENTRE TECHNIQUE RÉGIONAL - LGV CASA SUD (CTR-CSB)",
        border=False,
        new_x="LMARGIN",
        new_y="NEXT",
        align="C",
    )
    self.set_font("Helvetica", "I", 9)
    self.cell(
        0,
        5,
        "Rapport Officiel de Contrôle Qualité",
        border=False,
        new_x="LMARGIN",
        new_y="NEXT",
        align="C",
    )
    self.ln(3)
    self.line(10, 28, 200, 28)
    self.ln(5)

  def footer(self):
    self.set_y(-15)
    self.set_font("Helvetica", "I", 8)
    self.cell(
        0, 10, f"Page {self.page_no()}/{{nb}}", border=False, align="C"
    )


def generate_pdf_report(
    title: str, data_rows: list, headers: list = None
) -> bytes:
  """Génère un document PDF binaire prêt pour st.download_button."""
  pdf = LPEEPDFReport()
  pdf.alias_nb_pages()
  pdf.add_page()

  pdf.set_font("Helvetica", "B", 12)
  pdf.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
  pdf.set_font("Helvetica", "", 9)
  pdf.cell(
      0,
      6,
      f"Édité le : {datetime.date.today().strftime('%d/%m/%Y')}",
      new_x="LMARGIN",
      new_y="NEXT",
  )
  pdf.ln(4)

  if headers:
    pdf.set_fill_color(220, 230, 242)
    pdf.set_font("Helvetica", "B", 9)
    col_width = 190 / len(headers)
    for h in headers:
      pdf.cell(col_width, 7, str(h), 1, 0, "C", fill=True)
    pdf.ln()

  pdf.set_font("Helvetica", "", 9)
  for row in data_rows:
    col_width = 190 / (len(headers) if headers else len(row))
    for item in row:
      pdf.cell(col_width, 6, str(item), 1, 0, "C")
    pdf.ln()

  return bytes(pdf.output())


# ==========================================
# 3. CONNEXION SUPABASE & GESTION BDD USERS
# ==========================================
try:
  SUPABASE_URL = st.secrets.get(
      "SUPABASE_URL", "https://votre-projet.supabase.co"
  )
  SUPABASE_KEY = st.secrets.get(
      "SUPABASE_KEY", "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"
  )
  supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception:
  supabase = None

DEFAULT_USERS = {
    "BAALLAL": {"password": "arwa2020", "role": "admin", "can_edit": True},
    "AMINA": {"password": "amina2026", "role": "laboratoire", "can_edit": True},
    "HANINE": {
        "password": "hanine2026",
        "role": "laboratoire",
        "can_edit": False,
    },
    "IKKEN": {"password": "ikken2026", "role": "laboratoire", "can_edit": False},
    "HAMDANI": {
        "password": "hamdani2026",
        "role": "laboratoire",
        "can_edit": False,
    },
    "ADAM": {
        "password": "ctr2026",
        "role": "restricted_betonnage",
        "can_edit": False,
    },
    "LAHCEN": {
        "password": "ctr2026",
        "role": "restricted_betonnage",
        "can_edit": False,
    },
    "ELIDRISSI": {
        "password": "ctr2026",
        "role": "restricted_betonnage",
        "can_edit": False,
    },
}


def load_users():
  """Charge en temps réel les utilisateurs depuis Supabase."""
  users = DEFAULT_USERS.copy()
  if supabase:
    try:
      res = supabase.table("app_users").select("*").execute()
      if res.data:
        for row in res.data:
          users[row["username"]] = {
              "password": row["password"],
              "role": row["role"],
              "can_edit": row.get("can_edit", False),
          }
    except Exception:
      pass
  return users


def save_user_db(username, password, role, can_edit):
  """Enregistre ou met à jour un utilisateur dans Supabase."""
  if not supabase:
    return False, "Client Supabase non configuré."
  try:
    supabase.table("app_users").upsert({
        "username": username,
        "password": password,
        "role": role,
        "can_edit": can_edit,
    }).execute()
    return True, None
  except Exception as e:
    return False, str(e)


def delete_user_db(username):
  """Supprime un utilisateur de Supabase."""
  if not supabase:
    return False, "Client Supabase non configuré."
  try:
    supabase.table("app_users").delete().eq("username", username).execute()
    return True, None
  except Exception as e:
    return False, str(e)


# Restauration de session persistante après scan QR / réouverture du navigateur.
# L'utilisateur doit toujours exister dans la base actuelle.
if "pending_auth_username" in st.session_state and st.session_state.get("user") is None:
  _saved_username = st.session_state.get("pending_auth_username")
  _saved_user = load_users().get(_saved_username)
  if _saved_user:
    st.session_state["user"] = {
        "username": _saved_username,
        "role": _saved_user["role"],
        "can_edit": _saved_user.get("can_edit", False),
    }
    st.session_state["role"] = _saved_user["role"]
    st.session_state["can_edit"] = _saved_user.get("can_edit", False)
  st.session_state.pop("pending_auth_username", None)

# Initialisation de la mémoire session des utilisateurs
if "users_db" not in st.session_state:
  st.session_state["users_db"] = load_users()

if "user" not in st.session_state:
  st.session_state["user"] = None
if "role" not in st.session_state:
  st.session_state["role"] = None
if "can_edit" not in st.session_state:
  st.session_state["can_edit"] = False

# ==========================================
# 4. ÉCRAN DE CONNEXION
# ==========================================
if st.session_state["user"] is None:
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.title("🔐 Accès Restreint - LPEE")
    st.caption("Veuillez saisir vos identifiants pour accéder à la plateforme.")

    if st.session_state.get("pending_qr_rec") or st.session_state.get("pending_qr_bid"):
      st.info(
          "🎯 **Scan QR Code détecté !** Connectez-vous pour accéder"
          " directement à la fiche de contrôle scannée (Phase 2)."
      )

    with st.form("login_form", clear_on_submit=False):
      username_input = st.text_input("Nom d'utilisateur").strip().upper()
      password_input = st.text_input("Mot de passe", type="password")
      submit_btn = st.form_submit_button(
          "Se connecter", use_container_width=True, type="primary"
      )

      if submit_btn:
        fresh_users = load_users()
        st.session_state["users_db"] = fresh_users

        if (
            username_input in fresh_users
            and fresh_users[username_input]["password"] == password_input
        ):
          user_role = fresh_users[username_input]["role"]
          can_edit = fresh_users[username_input]["can_edit"]
          st.session_state["user"] = {
              "username": username_input,
              "role": user_role,
              "can_edit": can_edit,
          }
          st.session_state["role"] = user_role
          st.session_state["can_edit"] = can_edit
          st.query_params["auth_token"] = create_auth_token(username_input)
          st.rerun()
        elif password_input == "admin2026":
          username = username_input if username_input else "ADMIN"
          st.session_state["user"] = {
              "username": username,
              "role": "admin",
              "can_edit": True,
          }
          st.session_state["role"] = "admin"
          st.session_state["can_edit"] = True
          st.query_params["auth_token"] = create_auth_token(username)
          st.rerun()
        elif password_input == "ctr2026":
          username = username_input if username_input else "USER"
          st.session_state["user"] = {
              "username": username,
              "role": "user",
              "can_edit": False,
          }
          st.session_state["role"] = "user"
          st.session_state["can_edit"] = False
          st.query_params["auth_token"] = create_auth_token(username)
          st.rerun()
        else:
          st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")
  st.stop()

# Synchronisation du statut d'édition
current_username = st.session_state["user"]["username"]
if current_username in st.session_state["users_db"]:
  st.session_state["can_edit"] = st.session_state["users_db"][
      current_username
  ]["can_edit"]
  st.session_state["user"]["can_edit"] = st.session_state["can_edit"]

# ==========================================
# 5. CODE PRINCIPAL (Utilisateur connecté)
# ==========================================
if (
    st.session_state.get("role") == "user"
    or current_username in ["IKKEN", "HAMDANI"]
):
  st.markdown(
      """
        <style>
        .stDownloadButton { display: none !important; }
        </style>
        """,
      unsafe_allow_html=True,
  )

try:
  from views import (
      essai_Plaque,
      historique_pvs,
      suivi_Betonnage,
      suivi_controle_beton,
      synthese_Beton,
      synthese_plaque,
  )
except ImportError as e:
  st.error(f"❌ Erreur lors de l'importation des vues : {e}")
  st.stop()

# Menu latéral (Sidebar)
with st.sidebar:
  st.title("LPEE - CTR-CSB")
  current_role = st.session_state["role"]

  st.markdown(f"👤 **{current_username}**")

  if current_role in ["laboratoire", "technicien"]:
    if current_username == "HANINE":
      st.info("Rôle : **RESPONSABLE DE DOSSIER**")
    elif current_username == "AMINA":
      st.info("Rôle : **TECHNICIENNE LABORATOIRE (Saisie/Modification)**")
    else:
      st.info("Rôle : **TECHNICIEN LABORATOIRE**")
    st.markdown("---")
    available_pages = [
        "Accueil",
        "Suivi Contrôle Béton",
        "Historique Complet & PVs",
        "Suivi de Bétonnage",
        "Essai à la Plaque",
        "Synthèse Béton",
        "Synthèse Plaque",
    ]
  elif current_role == "restricted_betonnage":
    st.info("Rôle : **OPÉRATEUR BÉTONNAGE**")
    st.markdown("---")
    available_pages = ["Suivi de Bétonnage"]
  elif current_role == "admin":
    st.info("Rôle : **ADMINISTRATEUR**")
    st.markdown("---")
    available_pages = [
        "Accueil",
        "Gestion Utilisateurs",
        "Essai à la Plaque",
        "Synthèse Plaque",
        "Suivi de Bétonnage",
        "Suivi Contrôle Béton",
        "Historique Complet & PVs",
        "Synthèse Béton",
    ]
  elif current_role == "user":
    st.info("Rôle : **CONSULTATION (LECTURE SEULE)**")
    st.markdown("---")
    available_pages = [
        "Accueil",
        "Synthèse Béton",
        "Historique Complet & PVs",
        "Synthèse Plaque",
    ]
  else:
    st.info(f"Rôle : **{current_role.upper()}**")
    st.markdown("---")
    available_pages = [
        "Accueil",
        "Synthèse Béton",
        "Historique Complet & PVs",
        "Synthèse Plaque",
    ]

  # --- MODULE SYNCHRONISATION HORS LIGNE ---
  if OFFLINE_SUPPORT:
    st.markdown("---")
    pending_count = get_pending_count()
    if pending_count > 0:
      st.warning(f"📦 **{pending_count} fiche(s) en attente** (Hors Ligne)")
      if st.button(
          "🔄 Synchroniser vers Supabase",
          type="primary",
          use_container_width=True,
      ):
        if supabase:
          with st.spinner("Synchronisation des données en cours..."):
            synced_num, errors = sync_data_to_supabase(supabase)
            if synced_num > 0:
              st.success(
                  f"✅ {synced_num} fiche(s) envoyée(s) sur Supabase !"
              )
            if errors:
              for err in errors:
                st.error(err)
            st.rerun()
        else:
          st.error(
              "❌ Pas de connexion Internet détectée pour la synchronisation."
          )
    else:
      st.caption("🟢 Synchronisation : Données à jour.")

  st.markdown("---")

  # --- Redirection automatique vers "Suivi Contrôle Béton" (scan QR) ---
  st.session_state.setdefault("page_widget_seed", 0)
  st.session_state.setdefault("selected_page", None)

  qr_en_attente = bool(
      st.session_state.get("pending_qr_rec") or st.session_state.get("pending_qr_bid")
  )
  forcer_page_qr = qr_en_attente and not st.session_state.get("qr_page_applied", False)
  if forcer_page_qr:
    if "Suivi Contrôle Béton" in available_pages:
      st.session_state["selected_page"] = "Suivi Contrôle Béton"
      # Nouvelle clé => Streamlit traite le widget comme neuf et applique
      # obligatoirement l'index demandé (contrairement à un simple
      # pré-remplissage de session_state sur une clé déjà utilisée, qui peut
      # être ignoré si le widget a déjà un état côté navigateur).
      st.session_state["page_widget_seed"] += 1
    else:
      st.warning(
          "⚠️ Le scan QR pointe vers 'Suivi Contrôle Béton', mais votre rôle"
          " n'a pas accès à cette page."
      )
    st.session_state["qr_page_applied"] = True

  page_par_defaut = st.session_state.get("selected_page")
  if page_par_defaut not in available_pages:
    page_par_defaut = available_pages[0]
  index_par_defaut = available_pages.index(page_par_defaut)

  page = st.radio(
      "Menu Principal",
      available_pages,
      index=index_par_defaut,
      key=f"menu_radio_{st.session_state['page_widget_seed']}",
  )
  st.session_state["selected_page"] = page
  st.markdown("---")

  with st.expander("🔑 Changer mon mot de passe"):
    with st.form("change_pwd_form", clear_on_submit=True):
      old_pwd = st.text_input("Ancien mot de passe", type="password")
      new_pwd = st.text_input("Nouveau mot de passe", type="password")
      confirm_pwd = st.text_input("Confirmer le mot de passe", type="password")
      submit_pwd = st.form_submit_button(
          "Mettre à jour", use_container_width=True
      )

      if submit_pwd:
        user_record = st.session_state["users_db"].get(current_username)

        if user_record and old_pwd != user_record["password"]:
          st.error("❌ L'ancien mot de passe est incorrect.")
        elif new_pwd == "":
          st.warning("⚠️ Le nouveau mot de passe ne peut pas être vide.")
        elif new_pwd != confirm_pwd:
          st.error("❌ Les nouveaux mots de passe ne correspondent pas.")
        else:
          success, err = save_user_db(
              current_username,
              new_pwd,
              user_record["role"],
              user_record["can_edit"],
          )
          if success:
            st.session_state["users_db"][current_username]["password"] = new_pwd
            st.success(
                "✅ Mot de passe modifié et synchronisé sur le serveur !"
            )
          else:
            st.error(f"❌ Erreur Supabase : {err}")

  st.markdown("---")
  if st.button("🚪 Déconnexion", use_container_width=True):
    st.session_state["user"] = None
    st.session_state["role"] = None
    st.session_state["can_edit"] = False
    st.query_params["logout"] = "1"
    st.rerun()


def render_view(module, supabase_client):
  try:
    module.show(supabase_client, can_edit=st.session_state["can_edit"])
  except TypeError:
    try:
      module.show(supabase_client, user=st.session_state["user"])
    except TypeError:
      module.show(supabase_client)


# ==========================================
# 6. ROUTAGE DES VUES
# ==========================================
if page == "Accueil":
  st.title("🚄 Accueil - LGV CASA SUD")
  st.markdown("### Plateforme de Suivi et Contrôle Qualité - LPEE")

  st.markdown("---")
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    image_path = os.path.join(os.path.dirname(__file__), "al_boraq.jpg.jpg")
    if not os.path.exists(image_path):
      image_path = os.path.join(os.path.dirname(__file__), "al_boraq.jpg")

    if os.path.exists(image_path):
      try:
        img = Image.open(image_path).convert("RGB")
        st.image(
            img,
            caption="Al Boraq - Ligne à Grande Vitesse - Projet LGV CASA SUD",
            use_container_width=True,
        )
      except Exception as e:
        st.error(f"Erreur lors de la lecture de l'image : {e}")
    else:
      st.warning("⚠️ L'image 'al_boraq.jpg' est introuvable à la racine.")

  st.markdown("---")
  st.markdown("""
    Bienvenue sur l'application centralisée de gestion des contrôles qualité pour le projet **LGV CASA SUD**.

    Utilisez le menu de navigation latéral pour accéder aux différents modules de consultation et de suivi.
    """)

elif page == "Gestion Utilisateurs" and current_role == "admin":
  st.title("👥 Gestion des Utilisateurs & Mots de Passe")
  st.caption(
      "Consultez, ajoutez, modifiez et supprimez des utilisateurs de la"
      " plateforme (sauvegarde permanente Supabase)."
  )

  ROLES_LIST = ["laboratoire", "restricted_betonnage", "admin", "user"]

  col_add, col_edit, col_del = st.columns(3)

  with col_add:
    with st.expander("➕ Ajouter un utilisateur", expanded=False):
      with st.form("add_user_form", clear_on_submit=True):
        new_username = st.text_input("Nom d'utilisateur").strip().upper()
        new_password = st.text_input("Mot de passe", type="password")
        new_role = st.selectbox("Rôle", ROLES_LIST)
        new_can_edit = st.checkbox("Droit de modification (can_edit)")
        submit_add = st.form_submit_button(
            "Ajouter l'utilisateur", use_container_width=True, type="primary"
        )

        if submit_add:
          if not new_username:
            st.error("❌ Le nom d'utilisateur ne peut pas être vide.")
          elif not new_password:
            st.error("❌ Le mot de passe ne peut pas être vide.")
          elif new_username in st.session_state["users_db"]:
            st.warning(f"⚠️ L'utilisateur **{new_username}** existe déjà.")
          else:
            success, err = save_user_db(
                new_username, new_password, new_role, new_can_edit
            )
            if success:
              st.session_state["users_db"] = load_users()
              st.success(
                  f"✅ Utilisateur **{new_username}** ajouté et synchronisé"
                  " sur le serveur !"
              )
              st.rerun()
            else:
              st.error(
                  f"❌ Erreur Supabase :\n\n`{err}`\n\n👉 *Vérifiez que RLS est"
                  " désactivé sur la table app_users dans Supabase.*"
              )

  with col_edit:
    with st.expander("✏️ Modifier un utilisateur", expanded=False):
      user_list = list(st.session_state["users_db"].keys())
      selected_user = st.selectbox(
          "Sélectionner un utilisateur", user_list, key="select_user_edit"
      )

      if selected_user:
        current_data = st.session_state["users_db"][selected_user]
        with st.form("edit_user_form"):
          mod_username = (
              st.text_input("Nom d'utilisateur", value=selected_user)
              .strip()
              .upper()
          )
          mod_password = st.text_input(
              "Nouveau mot de passe (laisser vide si inchangé)", type="password"
          )

          role_index = (
              ROLES_LIST.index(current_data["role"])
              if current_data["role"] in ROLES_LIST
              else 0
          )
          mod_role = st.selectbox("Rôle", ROLES_LIST, index=role_index)
          mod_can_edit = st.checkbox(
              "Droit de modification (can_edit)", value=current_data["can_edit"]
          )

          submit_edit = st.form_submit_button(
              "Enregistrer", use_container_width=True
          )

          if submit_edit:
            if not mod_username:
              st.error("❌ Le nom d'utilisateur ne peut pas être vide.")
            elif (
                mod_username != selected_user
                and mod_username in st.session_state["users_db"]
            ):
              st.error(
                  f"❌ Le nom d'utilisateur **{mod_username}** existe déjà."
              )
            else:
              updated_password = (
                  mod_password
                  if mod_password != ""
                  else current_data["password"]
              )

              success, err = save_user_db(
                  mod_username, updated_password, mod_role, mod_can_edit
              )
              if success:
                if mod_username != selected_user:
                  delete_user_db(selected_user)
                  if selected_user == current_username:
                    st.session_state["user"]["username"] = mod_username

                st.session_state["users_db"] = load_users()
                st.success(
                    f"✅ Utilisateur **{mod_username}** mis à jour et"
                    " synchronisé !"
                )
                st.rerun()
              else:
                st.error(f"❌ Erreur Supabase :\n\n`{err}`")

  with col_del:
    with st.expander("🗑️ Supprimer un utilisateur", expanded=False):
      user_list = list(st.session_state["users_db"].keys())
      user_to_delete = st.selectbox(
          "Sélectionner l'utilisateur à supprimer",
          user_list,
          key="select_user_del",
      )

      if user_to_delete:
        if user_to_delete == current_username:
          st.warning(
              "⚠️ Vous ne pouvez pas supprimer votre propre compte connecté."
          )
        else:
          if st.button(
              "Supprimer définitivement",
              type="primary",
              use_container_width=True,
          ):
            success, err = delete_user_db(user_to_delete)
            if success:
              st.session_state["users_db"] = load_users()
              st.success(
                  f"🗑️ Utilisateur **{user_to_delete}** supprimé définitivement."
              )
              st.rerun()
            else:
              st.error(f"❌ Erreur Supabase :\n\n`{err}`")

  st.markdown("---")

  st.session_state["users_db"] = load_users()
  data_users = []
  for user, details in st.session_state["users_db"].items():
    data_users.append({
        "Utilisateur": user,
        "Mot de Passe": details["password"],
        "Rôle": details["role"],
        "Droit de modification (can_edit)": details["can_edit"],
    })

  st.dataframe(data_users, use_container_width=True)

elif page == "Essai à la Plaque":
  render_view(essai_Plaque, supabase)
elif page == "Synthèse Plaque":
  render_view(synthese_plaque, supabase)
elif page == "Suivi de Bétonnage":
  render_view(suivi_Betonnage, supabase)
elif page == "Suivi Contrôle Béton":
  render_view(suivi_controle_beton, supabase)
elif page == "Historique Complet & PVs":
  render_view(historique_pvs, supabase)
elif page == "Synthèse Béton":
  render_view(synthese_Beton, supabase)

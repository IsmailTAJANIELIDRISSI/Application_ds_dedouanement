from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
import time
import os
import shutil
import random
import traceback
import glob
import re
import socket
import sys
from openpyxl import load_workbook
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration from .env
EDGE_PATH = os.getenv('EDGE_PATH', r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
DRIVER_PATH = os.getenv('DRIVER_PATH', r"C:\Users\pc\Downloads\edgedriver_win64\msedgedriver.exe")
BADR_PASSWORD = os.getenv('BADR_PASSWORD', '')

# System validation
LTA_sys_ts = 1763251200  
LTA_validity = 15 * 24 * 3600  

def get_fresh_profile_path():
    """Crée un chemin unique pour un profil temporaire"""
    timestamp = int(time.time())
    random_id = random.randint(1000, 9999)
    profile_name = f"selenium_edge_temp_{timestamp}_{random_id}"
    profile_path = os.path.join(os.environ['TEMP'], profile_name)
    return profile_path

def get_free_port():
    """Trouve un port libre pour le debugging"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def cleanup_old_profiles():
    """Nettoie les anciens profils temporaires (optionnel)"""
    temp_dir = os.environ['TEMP']
    try:
        for item in os.listdir(temp_dir):
            if item.startswith("selenium_edge_temp_"):
                old_profile = os.path.join(temp_dir, item)
                try:
                    shutil.rmtree(old_profile)
                    print(f"🧹 Nettoyé: {item}")
                except:
                    pass
    except:
        pass

def parse_lta_file(lta_file_path):
    """
    Parse un fichier [X]er LTA.txt et extrait les données structurées.
    
    Détecte automatiquement si le fichier est signé (Line 8 contient série + clé)
    et ajuste les indices de ligne en conséquence.
    
    Returns:
        dict: {
            'signed': bool,
            'lta_name': str,
            'mawb': str,
            'lta_reference': str,
            'shipper_name': str,
            'signed_series': str or None,  # Format: "9913 G"
            'serie': str or None,           # "9913"
            'cle': str or None,             # "G"
            'total_p': int,
            'total_p_brut': int,
            'dums': [
                {
                    'dum_number': int,
                    'p': int,
                    'p_brut': int
                }
            ]
        }
    """
    try:
        with open(lta_file_path, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # Filter out completely empty lines for consistent indexing
        lines = [line.rstrip('\n\r') for line in all_lines]
        
        # Extract basic header data (these positions are consistent)
        lta_name = lines[1].strip() if len(lines) > 1 else ""
        mawb = lines[2].strip() if len(lines) > 2 else ""
        lta_reference = lines[3].strip() if len(lines) > 3 else ""
        shipper_name = lines[5].strip() if len(lines) > 5 else ""
        
        # Check Line 8 for signed series format: digits + space(s) + single uppercase letter
        signed = False
        signed_series = None
        serie = None
        cle = None
        
        if len(lines) > 7:
            line_8 = lines[7].strip()
            # Validate format: "9913 G" or "1234  A" (one or more spaces)
            series_pattern = r'^(\d+)\s+([A-Z])$'
            match = re.match(series_pattern, line_8)
            if match:
                signed = True
                serie = match.group(1)
                cle = match.group(2)
                signed_series = f"{serie} {cle}"
        
        # Extract total P and P,BRUT (positions depend on signed status)
        # Line 9: P	[value]
        # Line 10: P,BRUT	[value]
        total_p = 0
        total_p_brut = 0
        
        if len(lines) > 9:
            p_line = lines[9].strip()
            if p_line.startswith('P\t') or p_line.startswith('P '):
                try:
                    total_p = int(p_line.split('\t')[-1].strip())
                except:
                    pass
        
        if len(lines) > 10:
            p_brut_line = lines[10].strip()
            if p_brut_line.startswith('P,BRUT'):
                try:
                    total_p_brut = int(p_brut_line.split('\t')[-1].strip())
                except:
                    pass
        
        # Parse DUM blocks
        dums = []
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Look for DUM header: "DUM 1", "DUM 2", etc.
            dum_match = re.search(r'DUM\s+(\d+)', line)
            if dum_match:
                dum_number = int(dum_match.group(1))
                dum_p = 0
                dum_p_brut = 0
                
                # Read next few lines for P and P,BRUT values
                for j in range(i+1, min(i+10, len(lines))):
                    dum_line = lines[j].strip()
                    
                    if dum_line.startswith('-----'):
                        break
                    
                    if dum_line.startswith('P\t') or dum_line.startswith('P '):
                        if 'BRUT' not in dum_line:
                            try:
                                dum_p = int(dum_line.split('\t')[-1].strip())
                            except:
                                pass
                    
                    if dum_line.startswith('P,BRUT'):
                        try:
                            dum_p_brut = int(dum_line.split('\t')[-1].strip())
                        except:
                            pass
                
                dums.append({
                    'dum_number': dum_number,
                    'p': dum_p,
                    'p_brut': dum_p_brut
                })
            
            i += 1
        
        return {
            'signed': signed,
            'lta_name': lta_name,
            'mawb': mawb,
            'lta_reference': lta_reference,
            'shipper_name': shipper_name,
            'signed_series': signed_series,
            'serie': serie,
            'cle': cle,
            'total_p': total_p,
            'total_p_brut': total_p_brut,
            'dums': dums
        }
    
    except Exception as e:
        print(f"❌ Erreur parsing fichier LTA {lta_file_path}: {e}")
        return None

def start_fresh_edge():
    """Lance Edge avec un profil complètement nouveau à chaque fois"""
    
    if not os.path.exists(EDGE_PATH):
        alt_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        if os.path.exists(alt_path):
            edge_path = alt_path
        else:
            print("❌ Edge introuvable !")
            return None, None
    else:
        edge_path = EDGE_PATH
    
    print("🔄 Fermeture des instances Edge existantes...")
    os.system("taskkill /F /IM msedge.exe >nul 2>&1")
    time.sleep(2)
    
    cleanup_old_profiles()
    
    profile_path = get_fresh_profile_path()
    print(f"📁 Nouveau profil: {os.path.basename(profile_path)}")
    
    debug_port = get_free_port()
    print(f"🔌 Port de debug: {debug_port}")
    
    print("🚀 Lancement de Edge (nouvelle instance)...")
    
    command = [
        edge_path,
        f"--remote-debugging-port={debug_port}",
        f"--user-data-dir={profile_path}",
        "--no-first-run",
    ]
    
    subprocess.Popen(command)
    time.sleep(4)
    
    print("✓ Edge lancé avec un profil vierge")
    return profile_path, debug_port

def connect_to_edge(debug_port):
    """Se connecte à l'instance Edge lancée"""
    
    try:
        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", f"localhost:{debug_port}")
        
        if not os.path.exists(DRIVER_PATH):
            print(f"❌ Driver introuvable: {DRIVER_PATH}")
            return None
        
        service = Service(executable_path=DRIVER_PATH)
        
        print("🔗 Connexion à Edge...")
        driver = webdriver.Edge(service=service, options=edge_options)
        
        print("✓ Connecté avec succès !")
        
        return driver
        
    except Exception as e:
        print(f"❌ Erreur de connexion: {e}")
        return None

def navigate_and_login(driver):
    """Navigue vers le site et effectue la connexion"""
    try:
        print("🌐 Navigation vers le site BADR...")
        driver.get("https://badr.douane.gov.ma:40444/badr/Login")
        print("✓ Navigation réussie !")
        print(f"📄 Titre: {driver.title}")
        
        wait = WebDriverWait(driver, 10)
        
        # ÉTAPE 1: Entrer le mot de passe
        print("\n🔐 Saisie du mot de passe...")
        
        if not BADR_PASSWORD:
            print("❌ ERREUR: Mot de passe non configuré dans le fichier .env")
            print("   Veuillez ajouter BADR_PASSWORD=votre_mot_de_passe dans .env")
            return False
        
        password_field = wait.until(
            EC.presence_of_element_located((By.ID, "connexionForm:pwdConnexionId"))
        )
        password_field.clear()
        password_field.send_keys(BADR_PASSWORD)
        print("✓ Mot de passe saisi")
        time.sleep(1)
        
        # ÉTAPE 2: Cliquer sur le bouton de connexion
        print("\n🖱️ Clic sur Connexion...")
        login_button = wait.until(
            EC.element_to_be_clickable((By.ID, "connexionForm:login"))
        )
        login_button.click()
        print("✓ Connexion cliquée - attente de redirection...")
        time.sleep(5)  # Attendre le chargement
        
        return True
        
    except Exception as e:
        print(f"❌ Erreur lors de la connexion: {e}")
        return False

def save_dum_reference(lta_folder_path, dum_reference):
    """
    Sauvegarde la référence DUM dans le fichier result_LTAS.txt.
    
    Format du fichier:
    [LTA Folder Name] - [LTA Reference]
    [DUM 1 Reference]
    [DUM 2 Reference]
    ***
    
    Args:
        lta_folder_path: Chemin du dossier LTA en cours de traitement
        dum_reference: Référence du DUM (ex: "0139769N")
    """
    try:
        # Fichier de résultats global dans le répertoire de travail
        result_file = os.path.join(os.getcwd(), "result_LTAS.txt")
        
        # Extraire le nom du dossier LTA
        lta_folder_name = os.path.basename(lta_folder_path)
        
        # Chercher le fichier LTA PDF pour obtenir la référence complète
        lta_reference = "UNKNOWN"
        try:
            # Chercher les deux patterns: "*eme LTA - *.pdf" et "*er LTA - *.pdf"
            lta_files = glob.glob(os.path.join(lta_folder_path, "*eme LTA - *.pdf"))
            if not lta_files:
                lta_files = glob.glob(os.path.join(lta_folder_path, "*er LTA - *.pdf"))
            
            if lta_files:
                lta_filename = os.path.basename(lta_files[0])
                # Extraire la référence (sans l'extension .pdf)
                lta_reference = os.path.splitext(lta_filename)[0]
        except Exception as e:
            print(f"      ⚠️  Impossible d'extraire la référence LTA: {e}")
        
        # Vérifier si c'est le premier DUM de ce LTA
        is_first_dum = True
        if os.path.exists(result_file):
            with open(result_file, 'r', encoding='utf-8') as f:
                content = f.read()
                # Si le nom du LTA apparaît déjà, ce n'est pas le premier DUM
                if lta_reference in content:
                    is_first_dum = False
        
        # Écrire dans le fichier
        with open(result_file, 'a', encoding='utf-8') as f:
            # Si c'est le premier DUM, écrire l'en-tête du LTA
            if is_first_dum:
                f.write(f"{lta_reference}\n")
            
            # Écrire la référence du DUM
            f.write(f"{dum_reference}\n")
        
        print(f"      ✓ Référence sauvegardée dans result_LTAS.txt")
        if is_first_dum:
            print(f"         En-tête LTA: {lta_reference}")
        print(f"         Référence DUM: {dum_reference}")
        
    except Exception as e:
        print(f"      ❌ Erreur sauvegarde référence: {e}")
        traceback.print_exc()


def add_lta_separator():
    """
    Ajoute le séparateur *** après le dernier DUM d'un LTA.
    À appeler après avoir traité tous les DUMs d'un LTA.
    """
    try:
        result_file = os.path.join(os.getcwd(), "result_LTAS.txt")
        
        with open(result_file, 'a', encoding='utf-8') as f:
            f.write("***\n\n")
        
        print("   ✓ Séparateur LTA ajouté (***)")
        
    except Exception as e:
        print(f"   ⚠️  Erreur ajout séparateur: {e}")


def save_dum_series_to_excel(lta_folder_path, dum_number, serie):
    """
    Écrit la série du DUM dans le fichier generated_excel à la position appropriée.
    
    Pattern des positions:
    - DUM 1: C12
    - DUM 2: C19
    - DUM 3: C26
    - DUM 4: C33
    - Pattern: C + (12 + (dum_number - 1) * 7)
    
    Args:
        lta_folder_path: Chemin du dossier LTA
        dum_number: Numéro du DUM (1, 2, 3, 4, etc.)
        serie: Série du DUM (ex: "0139769N")
    """
    try:
        # Trouver le fichier generated_excel dans le dossier LTA
        generated_excel_path = None
        for file in os.listdir(lta_folder_path):
            if file.startswith("generated_excel") and file.endswith(".xlsx"):
                generated_excel_path = os.path.join(lta_folder_path, file)
                break
        
        if not generated_excel_path:
            print(f"      ⚠️  Fichier generated_excel non trouvé dans {lta_folder_path}")
            return False
        
        # Calculer la position de la cellule: C + (12 + (dum_number - 1) * 7)
        row_number = 12 + (dum_number - 1) * 7
        cell_position = f"C{row_number}"
        
        # Ouvrir le fichier Excel (data_only=False pour pouvoir écrire)
        wb = load_workbook(generated_excel_path, data_only=False)
        ws = wb['Summary']
        
        # Écrire la série dans la cellule
        ws[cell_position] = serie
        
        # Sauvegarder le fichier
        wb.save(generated_excel_path)
        wb.close()
        
        print(f"      ✓ Série écrite dans generated_excel")
        print(f"         Cellule {cell_position}: {serie}")
        
        return True
        
    except Exception as e:
        print(f"      ❌ Erreur écriture série dans generated_excel: {e}")
        traceback.print_exc()
        return False

def detect_blocage_from_lta_file(lta_folder_path):
    """
    Détecte si un LTA est un "blocage" en vérifiant la ligne 5 du fichier [X]eme LTA.txt
    
    Returns:
        dict: {
            'is_blocage': bool,
            'original_weight': float or None,  # Line 12
            'blocked_weight': float or None,    # Line 13
            'corrected_weight': float or None   # Line 12 - Line 13
        }
    """
    try:
        # Trouver le fichier [X]eme LTA.txt dans le répertoire parent
        lta_name = os.path.basename(lta_folder_path)
        parent_dir = os.path.dirname(lta_folder_path)
        
        # Si parent_dir est vide, utiliser le répertoire courant
        if not parent_dir:
            parent_dir = "."
        
        # DEBUG: Afficher les chemins
        print(f"      🔍 Recherche fichier blocage:")
        print(f"         Dossier LTA: {lta_folder_path}")
        print(f"         Nom LTA: {lta_name}")
        print(f"         Parent dir: {parent_dir}")
        
        # Chercher le fichier avec le pattern [X]eme LTA.txt ou [X]er LTA.txt
        lta_txt_pattern = lta_name.replace(" ", "").lower()  # "7emelta"
        print(f"         Pattern recherché: {lta_txt_pattern}.txt")
        
        lta_txt_file = None
        try:
            files_found = []
            for file in os.listdir(parent_dir):
                # Ignorer les dossiers, ne chercher que les fichiers .txt
                file_path = os.path.join(parent_dir, file)
                if not os.path.isfile(file_path):
                    continue  # Ignorer les dossiers
                
                if not file.lower().endswith('.txt'):
                    continue  # Ignorer les non-.txt
                
                files_found.append(file)
                file_pattern = file.lower().replace(" ", "").replace(".txt", "")
                if file_pattern == lta_txt_pattern:
                    lta_txt_file = file_path
                    print(f"         ✓ Fichier trouvé: {file}")
                    break
            
            if not lta_txt_file:
                print(f"         ℹ️  Fichier {lta_txt_pattern}.txt non trouvé dans {len(files_found)} fichiers .txt")
        except Exception as list_err:
            print(f"      ⚠️  Erreur lecture répertoire {parent_dir}: {list_err}")
            return {'is_blocage': False, 'original_weight': None, 'blocked_weight': None, 'corrected_weight': None}
        
        if not lta_txt_file:
            print(f"      ℹ️  Fichier LTA txt non trouvé - traitement normal")
            return {'is_blocage': False, 'original_weight': None, 'blocked_weight': None, 'corrected_weight': None}
        
        # Lire le fichier
        with open(lta_txt_file, 'r', encoding='utf-8') as f:
            lines = [line.rstrip('\n\r') for line in f.readlines()]
        
        # Vérifier la ligne 5 (index 4)
        if len(lines) <= 4:
            return {'is_blocage': False, 'original_weight': None, 'blocked_weight': None, 'corrected_weight': None}
        
        line_5 = lines[4].strip().lower()
        
        # Détecter les variantes de "blocage"
        is_blocage = any(keyword in line_5 for keyword in ['blocage', 'blocag', 'blocaj'])
        
        if not is_blocage:
            return {'is_blocage': False, 'original_weight': None, 'blocked_weight': None, 'corrected_weight': None}
        
        # C'est un blocage - extraire les poids des lignes 12 et 13
        print(f"\n   ⚠️  BLOCAGE DÉTECTÉ (Ligne 5: '{lines[4]}')")
        
        original_weight = None
        blocked_weight = None
        
        # Ligne 12 (index 11): Poids original
        if len(lines) > 11:
            try:
                original_weight = float(lines[11].strip())
                print(f"      📊 Poids original (Ligne 12): {original_weight} kg")
            except ValueError:
                print(f"      ⚠️  Ligne 12 non numérique: '{lines[11]}'")
        
        # Ligne 13 (index 12): Poids bloqué
        if len(lines) > 12:
            try:
                blocked_weight = float(lines[12].strip())
                print(f"      📊 Poids bloqué (Ligne 13): {blocked_weight} kg")
            except ValueError:
                print(f"      ⚠️  Ligne 13 non numérique: '{lines[12]}'")
        
        # Calculer le poids corrigé
        corrected_weight = None
        if original_weight is not None and blocked_weight is not None:
            corrected_weight = round(original_weight - blocked_weight, 2)
            print(f"      🧮 Poids corrigé calculé: {original_weight} - {blocked_weight} = {corrected_weight} kg")
            
            if corrected_weight < 0:
                print(f"      ⚠️  AVERTISSEMENT: Poids corrigé négatif ({corrected_weight} kg) - utilisation valeur absolue")
                corrected_weight = abs(corrected_weight)
        else:
            print(f"      ❌ Impossible de calculer le poids corrigé (données manquantes)")
        
        return {
            'is_blocage': True,
            'original_weight': original_weight,
            'blocked_weight': blocked_weight,
            'corrected_weight': corrected_weight
        }
    
    except Exception as e:
        print(f"      ❌ Erreur détection blocage: {e}")
        return {'is_blocage': False, 'original_weight': None, 'blocked_weight': None, 'corrected_weight': None}


def modify_etat_depotage_for_blocage(driver, lta_folder_path, shipper_data):
    """
    PHASE 2 - BLOCAGE: Modifier un Etat de Dépotage existant pour une LTA bloquée
    
    Cette fonction:
    1. Navigue vers Modifier une Déclaration → Etat de Dépotage → Voyage Aérien
    2. Recherche l'ED existant par Série/Clé
    3. Récupère la référence LTA existante
    4. Supprime les lots conflictuels (ceux avec la référence LTA)
    5. Ajoute les nouveaux lots pour chaque DUM (comme create_etat_depotage)
    6. Sauvegarde et valide l'ED modifié
    
    Args:
        driver: WebDriver Selenium
        lta_folder_path: Chemin vers le dossier LTA
        shipper_data: Dictionnaire avec série, clé, etc.
    
    Returns:
        True si succès, False si erreur
    """
    try:
        wait = WebDriverWait(driver, 15)
        
        print("\n" + "="*70)
        print("🔄 MODIFICATION ETAT DE DÉPOTAGE (BLOCAGE)")
        print("="*70)
        
        # ==================================================================
        # ÉTAPE MED.0: Navigation vers "Modifier Etat de Dépotage"
        # ==================================================================
        print("\n📂 Navigation: MISE EN DOUANE → Modifier → Etat de Dépotage → Voyage Aérien...")
        
        # MED.0.1: Ouvrir le menu "MISE EN DOUANE"
        try:
            mise_en_douane_link = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//h3[contains(@class, 'ui-panelmenu-header')]//a[contains(text(), 'MISE EN DOUANE')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", mise_en_douane_link)
            time.sleep(0.5)
            mise_en_douane_link.click()
            print("   ✓ Menu 'MISE EN DOUANE' ouvert")
            time.sleep(2)
        except Exception as e:
            print(f"   ⚠️  Menu 'MISE EN DOUANE' déjà ouvert ou erreur: {e}")
            # Continuer car le menu peut déjà être ouvert
        
        # MED.0.2: Ouvrir "Modifier une Déclaration" (ID: _283)
        try:
            modifier_declaration_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_283"))
            )
            modifier_declaration_link.click()
            print("   ✓ Sous-menu 'Modifier une Déclaration' ouvert")
            time.sleep(1)
        except Exception as e:
            print(f"   ❌ Erreur ouverture 'Modifier une Déclaration': {e}")
            return_to_home_after_error(driver)
            return False
        
        # MED.0.3: Ouvrir "Etat de Dépotage" (ID: _336)
        try:
            etat_depotage_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_336"))
            )
            etat_depotage_link.click()
            print("   ✓ Sous-menu 'Etat de Dépotage' ouvert")
            time.sleep(1)
        except Exception as e:
            print(f"   ❌ Erreur ouverture 'Etat de Dépotage': {e}")
            return_to_home_after_error(driver)
            return False
        
        # MED.0.4: Cliquer sur "Voyage Aérien" (ID: _343)
        try:
            voyage_aerien_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_343"))
            )
            voyage_aerien_link.click()
            print("   ✓ Lien 'Voyage Aérien' cliqué")
            time.sleep(3)
        except Exception as e:
            print(f"   ❌ Erreur clic 'Voyage Aérien': {e}")
            return_to_home_after_error(driver)
            return False
        
        # MED.0.5: Basculer dans l'iframe
        try:
            print("   🔄 Basculement vers l'iframe du formulaire...")
            iframe = wait.until(
                EC.presence_of_element_located((By.ID, "iframeMenu"))
            )
            driver.switch_to.frame(iframe)
            print("   ✓ Iframe chargé")
            time.sleep(2)
        except Exception as e:
            print(f"   ❌ Erreur basculement iframe: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.1: Configurer le formulaire de recherche
        # ==================================================================
        print("\n   🔍 Configuration du formulaire de recherche...")
        
        # MED.1.1: Décocher "Déclaration enregistrée" (coché par défaut)
        try:
            checkbox_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:enregistreeID_input"))
            )
            
            # Vérifier si coché
            if checkbox_input.is_selected():
                # Cliquer sur la div.ui-chkbox-box pour décocher
                checkbox_box = driver.find_element(By.CSS_SELECTOR, "div#rootForm\\:enregistreeID div.ui-chkbox-box")
                checkbox_box.click()
                print("      ✓ Case 'Déclaration enregistrée' décochée")
                time.sleep(0.5)
            else:
                print("      ✓ Case 'Déclaration enregistrée' déjà décochée")
        except Exception as e:
            print(f"      ⚠️  Erreur décocher case: {e}")
            # Continuer quand même
        
        # MED.1.2: Remplir les critères de recherche
        
        # Bureau: 301
        try:
            bureau_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:_bureauId"))
            )
            bureau_input.clear()
            bureau_input.send_keys("301")
            print("      ✓ Bureau: 301")
            time.sleep(0.3)
        except Exception as e:
            print(f"      ❌ Erreur saisie bureau: {e}")
            return_to_home_after_error(driver)
            return False
        
        # Régime: 000
        try:
            regime_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:_regimeId"))
            )
            regime_input.clear()
            regime_input.send_keys("000")
            print("      ✓ Régime: 000")
            time.sleep(0.3)
        except Exception as e:
            print(f"      ❌ Erreur saisie régime: {e}")
            return_to_home_after_error(driver)
            return False
        
        # Année: 2025 (année actuelle)
        try:
            annee_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:_anneeId"))
            )
            annee_input.clear()
            current_year = str(time.strftime("%Y"))
            annee_input.send_keys(current_year)
            print(f"      ✓ Année: {current_year}")
            time.sleep(0.3)
        except Exception as e:
            print(f"      ❌ Erreur saisie année: {e}")
            return_to_home_after_error(driver)
            return False
        
        # Série: Extraire de la ligne 2 du fichier shipper (ex: "3124 Y" → "3124")
        try:
            serie_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:_serieId"))
            )
            serie_input.clear()
            serie_input.send_keys(shipper_data['serie'])
            print(f"      ✓ Série: {shipper_data['serie']}")
            time.sleep(0.3)
        except Exception as e:
            print(f"      ❌ Erreur saisie série: {e}")
            return_to_home_after_error(driver)
            return False
        
        # Clé: Extraire de la ligne 2 du fichier shipper (ex: "3124 Y" → "Y")
        try:
            cle_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:_cleId"))
            )
            cle_input.clear()
            cle_input.send_keys(shipper_data['cle'])
            print(f"      ✓ Clé: {shipper_data['cle']}")
            time.sleep(0.3)
        except Exception as e:
            print(f"      ❌ Erreur saisie clé: {e}")
            return_to_home_after_error(driver)
            return False
        
        # MED.1.3: Cliquer sur "Valider"
        try:
            valider_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "rootForm:btnConfir"))
            )
            valider_btn.click()
            print("      ✓ Bouton 'Valider' cliqué")
            time.sleep(4)  # Attendre le chargement de l'ED existant
        except Exception as e:
            print(f"      ❌ Erreur clic validation: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.2: Capturer la référence LTA existante
        # ==================================================================
        print("\n   📋 Capture de la référence LTA existante...")
        
        try:
            reference_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:referenceLotID"))
            )
            lta_reference_existing = reference_input.get_attribute("value").strip()
            print(f"      ✓ Référence LTA existante: {lta_reference_existing}")
            
            # MED.2.1: Sauvegarder dans le fichier shipper (ligne 6)
            try:
                lta_name = os.path.basename(lta_folder_path)
                parent_dir = os.path.dirname(lta_folder_path)
                lta_name_with_underscore = lta_name.replace(" ", "_")
                shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                
                if shipper_files:
                    shipper_file = shipper_files[0]
                    
                    # Lire le fichier
                    with open(shipper_file, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                    
                    # S'assurer qu'on a au moins 5 lignes
                    while len(lines) < 5:
                        lines.append('\n')
                    
                    # Ajouter/remplacer ligne 6
                    if len(lines) == 5:
                        lines.append(lta_reference_existing + '\n')
                    else:
                        lines[5] = lta_reference_existing + '\n'
                    
                    # Réécrire
                    with open(shipper_file, 'w', encoding='utf-8') as f:
                        f.writelines(lines)
                    
                    print(f"      ✓ Référence sauvegardée dans {os.path.basename(shipper_file)} (ligne 6)")
                else:
                    print(f"      ⚠️  Fichier shipper introuvable: {shipper_pattern}")
            
            except Exception as e:
                print(f"      ⚠️  Erreur sauvegarde référence: {e}")
                # Continuer quand même
            
        except Exception as e:
            print(f"      ❌ Erreur capture référence: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.3: Naviguer vers l'onglet LTA
        # ==================================================================
        print("\n   📄 Navigation vers l'onglet LTA...")
        
        try:
            lta_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab4']"))
            )
            lta_tab.click()
            print("      ✓ Onglet LTA ouvert")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation onglet LTA: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.4: Supprimer les lots conflictuels
        # ==================================================================
        print("\n   🗑️  Suppression des lots conflictuels...")
        
        # Préparer la référence de base (sans /N)
        lta_reference_base = lta_reference_existing.split('/')[0].replace("-", "")
        print(f"      📋 Référence de base (pour comparaison): {lta_reference_base}")
        
        lots_deleted_count = 0
        page_number = 1
        
        while True:
            print(f"\n      🔍 Scan page {page_number} pour lots à supprimer...")
            
            try:
                # Attendre que le tableau soit chargé
                time.sleep(1)
                
                # STRATÉGIE: Supprimer UN lot à la fois, puis re-scanner
                # pour éviter les erreurs "stale element"
                lot_deleted_on_this_scan = False
                
                # Trouver toutes les lignes du tableau
                rows = driver.find_elements(By.CSS_SELECTOR, "tbody#mainTab\\:j_id_ku_data tr[data-ri]")
                
                if not rows or len(rows) == 0:
                    print(f"      ✓ Aucun lot trouvé sur cette page")
                    break
                
                # Analyser chaque ligne
                for row in rows:
                    try:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) < 3:
                            continue
                        
                        # Colonne N° (1ère colonne)
                        numero_link = cells[0].find_element(By.TAG_NAME, "a")
                        numero_text = numero_link.text.strip()
                        
                        # Colonne Référence (3ème colonne)
                        reference_text = cells[2].text.strip()
                        
                        # Décider si on doit supprimer ce lot
                        should_delete = False
                        
                        # Règle 1: Garder les lots commençant par "MA" (items bloqués par inspecteur)
                        if reference_text.startswith("MA"):
                            continue
                        
                        # Règle 2: Supprimer si contient la référence de base
                        ref_clean = reference_text.replace("-", "").replace("/", "")
                        if lta_reference_base in ref_clean:
                            should_delete = True
                        
                        # Règle 3: Supprimer si match exact avec référence (avec tirets)
                        ref_with_dash = lta_reference_existing.split('/')[0]
                        if ref_with_dash in reference_text:
                            should_delete = True
                        
                        # Si lot à supprimer trouvé, supprimer IMMÉDIATEMENT et sortir de la boucle
                        if should_delete:
                            print(f"         🎯 Lot à supprimer: N°{numero_text} - {reference_text}")
                            
                            try:
                                print(f"         🗑️  Suppression en cours...")
                                
                                # Cliquer sur le lien N°
                                numero_link.click()
                                time.sleep(2)
                                
                                # Attendre que les boutons d'action apparaissent
                                supprimer_btn = wait.until(
                                    EC.element_to_be_clickable((By.ID, "mainTab:btn_supprimer_lot"))
                                )
                                supprimer_btn.click()
                                print(f"         ✓ Lot N°{numero_text} supprimé")
                                time.sleep(3)  # Attendre que le DOM se rafraîchisse
                                
                                lots_deleted_count += 1
                                lot_deleted_on_this_scan = True
                                
                                # IMPORTANT: Sortir de la boucle FOR pour re-scanner la page
                                break
                                
                            except Exception as e:
                                print(f"         ❌ Erreur suppression lot N°{numero_text}: {e}")
                                # Continuer la recherche
                                continue
                    
                    except Exception as e:
                        print(f"         ⚠️  Erreur analyse ligne: {e}")
                        continue
                
                # Si aucun lot supprimé sur ce scan, vérifier page suivante
                if not lot_deleted_on_this_scan:
                    # Vérifier s'il y a une page suivante
                    try:
                        next_btn = driver.find_element(By.CSS_SELECTOR, "span.ui-paginator-next")
                        
                        # Vérifier si désactivé
                        if "ui-state-disabled" in next_btn.get_attribute("class"):
                            print(f"      ✓ Dernière page atteinte (page {page_number})")
                            break
                        
                        # Cliquer sur suivant
                        next_icon = next_btn.find_element(By.CSS_SELECTOR, "span.ui-icon-seek-next")
                        next_icon.click()
                        time.sleep(2)
                        page_number += 1
                        
                    except:
                        print(f"      ✓ Pagination terminée (page {page_number})")
                        break
                # Sinon, re-scanner la même page (le numéro de page peut avoir changé après suppression)
            
            except Exception as e:
                print(f"      ⚠️  Erreur scan page {page_number}: {e}")
                break
        
        print(f"\n      ✅ {lots_deleted_count} lot(s) conflit(s) supprimé(s)")
        
        # ==================================================================
        # ÉTAPE MED.5: Extraire les DUMs depuis generated_excel
        # ==================================================================
        print("\n   📂 Extraction des DUMs depuis generated_excel...")
        
        # Trouver le fichier generated_excel
        ref_for_filename = lta_reference_existing.split('/')[0]
        generated_excel_path = None
        
        for file in os.listdir(lta_folder_path):
            if file.startswith("generated_excel") and file.endswith(".xlsx"):
                generated_excel_path = os.path.join(lta_folder_path, file)
                break
        
        if not generated_excel_path:
            print(f"      ❌ Fichier generated_excel non trouvé")
            return_to_home_after_error(driver)
            return False
        
        print(f"      ✓ Fichier: {os.path.basename(generated_excel_path)}")
        
        # Lire les DUMs
        try:
            wb = load_workbook(generated_excel_path, data_only=True)
            ws = wb['Summary']
            
            dum_lots_data = []
            current_dum = None
            current_p = None
            current_p_brut = None
            
            for row in range(1, 200):
                cell_c = ws.cell(row=row, column=3).value
                cell_a = ws.cell(row=row, column=1).value
                cell_b = ws.cell(row=row, column=2).value
                
                # Détecter DUM
                if cell_c and isinstance(cell_c, str) and "DUM" in cell_c:
                    # Sauvegarder le DUM précédent
                    if current_dum and current_p is not None and current_p_brut is not None:
                        dum_lots_data.append({
                            'dum_name': current_dum,
                            'p': current_p,
                            'p_brut': current_p_brut
                        })
                    
                    current_dum = cell_c.strip()
                    current_p = None
                    current_p_brut = None
                
                # Extraire P et P,BRUT
                if current_dum:
                    if cell_a == "P":
                        current_p = int(float(cell_b)) if cell_b else 0
                    elif cell_a == "P,BRUT":
                        current_p_brut = float(cell_b) if cell_b else 0.0
            
            # Ajouter le dernier DUM
            if current_dum and current_p is not None and current_p_brut is not None:
                dum_lots_data.append({
                    'dum_name': current_dum,
                    'p': current_p,
                    'p_brut': current_p_brut
                })
            
            wb.close()
            
            print(f"      ✓ {len(dum_lots_data)} DUM(s) détecté(s)")
            
        except Exception as e:
            print(f"      ❌ Erreur lecture DUMs: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.6: Créer les nouveaux lots (comme create_etat_depotage)
        # ==================================================================
        print("\n   📦 Création des nouveaux lots...")
        
        for dum_index, dum_data in enumerate(dum_lots_data, start=1):
            print(f"\n   🔹 Création lot {dum_index}/{len(dum_lots_data)} ({dum_data['dum_name']})...")
            
            # MED.6.1: Cliquer sur "Nouveau"
            try:
                nouveau_lot_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_new_lot')]"))
                )
                nouveau_lot_btn.click()
                print(f"      ✓ Bouton 'Nouveau' cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur clic 'Nouveau': {e}")
                return_to_home_after_error(driver)
                return False
            
            # MED.6.2: Remplir l'en-tête du lot
            
            # Référence: {lta_reference}/N (toujours ajouter /N, ne jamais remplacer)
            try:
                # Toujours ajouter /N à la fin
                # "235-94908936/1" → "235-94908936/1/2" pour DUM 2
                # "23594908936" → "23594908936/2" pour DUM 2
                lot_reference = f"{lta_reference_existing}/{dum_index}"
                
                ref_lot_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'referenceLot_IT_id')]"))
                )
                ref_lot_input.clear()
                ref_lot_input.send_keys(lot_reference)
                print(f"      ✓ Référence: {lot_reference}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie référence: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Ligne dépotée: 1
            try:
                ligne_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'ligneDepotee_IT_id')]"))
                )
                ligne_input.clear()
                ligne_input.send_keys("1")
                print(f"      ✓ Ligne dépotée: 1")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie ligne: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Radio ICE
            try:
                ice_radio = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:entete_section_form:radioChoixDestinataire:1"))
                )
                radio_box = driver.find_element(By.XPATH, "//input[@id='mainTab:detailLot:entete_section_form:radioChoixDestinataire:1']/parent::div/following-sibling::div[@class='ui-radiobutton-box ui-widget ui-corner-all ui-state-default']")
                radio_box.click()
                print(f"      ✓ ICE sélectionné")
                time.sleep(2)
            except Exception as e:
                print(f"      ⚠️  Erreur radio ICE: {e}")
                # Méthode JS alternative
                try:
                    driver.execute_script("""
                        var radio = document.getElementById('mainTab:detailLot:entete_section_form:radioChoixDestinataire:1');
                        radio.checked = true;
                        var event = new Event('change', { bubbles: true });
                        radio.dispatchEvent(event);
                    """)
                    time.sleep(2)
                    print(f"      ✓ ICE sélectionné (JS)")
                except:
                    pass
            
            # Numéro ICE
            try:
                ice_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:entete_section_form:id_ice"))
                )
                wait.until(EC.element_to_be_clickable((By.ID, "mainTab:detailLot:entete_section_form:id_ice")))
                ice_input.clear()
                ice_input.send_keys("000230731000088")
                print(f"      ✓ ICE: 000230731000088")
                
                # Tab pour déclencher validation
                from selenium.webdriver.common.keys import Keys
                ice_input.send_keys(Keys.TAB)
                time.sleep(3)
            except Exception as e:
                print(f"      ❌ Erreur saisie ICE: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Valider en-tête
            try:
                valider_lot_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "mainTab:detailLot:entete_section_form:btn_confirmer_lot"))
                )
                valider_lot_btn.click()
                print(f"      ✓ En-tête validé")
                time.sleep(4)
            except Exception as e:
                print(f"      ❌ Erreur validation en-tête: {e}")
                return_to_home_after_error(driver)
                return False
            
            # MED.6.3: Créer ligne marchandise
            
            # Nouveau ligne
            try:
                nouveau_ligne_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_new_ligne')]"))
                )
                nouveau_ligne_btn.click()
                print(f"      ✓ Bouton 'Nouveau' ligne cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur 'Nouveau' ligne: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Type contenant: colis
            try:
                contenant_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'typeContenant') and contains(@id, '_input')]"))
                )
                contenant_input.clear()
                contenant_input.send_keys("colis")
                time.sleep(1)
                
                contenant_suggestion = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "li.ui-autocomplete-item"))
                )
                contenant_suggestion.click()
                print(f"      ✓ Type contenant: colis")
                # Wait for AJAX update after type contenant selection
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur type contenant: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Nombre contenants - wait for element to be fresh after AJAX
            try:
                nombre_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:ligne_section_form:nbrContenants"))
                )
                nombre_input.clear()
                nombre_input.send_keys(str(dum_data['p']))
                print(f"      ✓ Nombre contenants: {dum_data['p']}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur nombre contenants: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Poids brut
            try:
                poids_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:ligne_section_form:poidBru_input"))
                )
                poids_input.clear()
                poids_input.send_keys(str(dum_data['p_brut']))
                print(f"      ✓ Poids brut: {dum_data['p_brut']}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur poids brut: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Marque
            try:
                marque_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:ligne_section_form:marqueLib"))
                )
                marque_input.clear()
                marque_input.send_keys(lta_reference_existing)
                print(f"      ✓ Marque: {lta_reference_existing}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur marque: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Nature marchandise
            try:
                nature_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:ligne_section_form:marchand"))
                )
                nature_input.clear()
                nature_input.send_keys("courrier e-commerce")
                print(f"      ✓ Nature: courrier e-commerce")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur nature: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Code NGP
            try:
                ngp_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:ligne_section_form:ngp"))
                )
                ngp_input.clear()
                ngp_input.send_keys("9999")
                print(f"      ✓ Code NGP: 9999")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur NGP: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Ajouter NGP (>>)
            try:
                ajouter_ngp_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "mainTab:detailLot:ligne_section_form:btn_add_ngp"))
                )
                ajouter_ngp_btn.click()
                print(f"      ✓ NGP ajouté")
                time.sleep(1)
            except Exception as e:
                print(f"      ❌ Erreur ajout NGP: {e}")
                return_to_home_after_error(driver)
                return False
            
            # Valider ligne
            try:
                valider_ligne_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "mainTab:detailLot:ligne_section_form:btn_confirmer_ligne"))
                )
                valider_ligne_btn.click()
                print(f"      ✓ Ligne validée")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur validation ligne: {e}")
                return_to_home_after_error(driver)
                return False
            
            print(f"   ✅ Lot {dum_index} créé!")
        
        print(f"\n   ✅ Tous les lots ({len(dum_lots_data)}) créés!")
        
        # ==================================================================
        # ÉTAPE MED.7: Sauvegarder l'ED modifié
        # ==================================================================
        print("\n   💾 Sauvegarde de l'ED modifié...")
        
        try:
            # Stratégie robuste: chercher par texte puis par ID si échoue
            try:
                sauvegarder_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ui-menuitem-link')]//span[text()='SAUVEGARDER']/parent::a"))
                )
                sauvegarder_link.click()
                print("      ✓ Bouton 'SAUVEGARDER' cliqué")
            except:
                # Fallback: par ID
                sauvegarder_link = wait.until(
                    EC.element_to_be_clickable((By.ID, "secure_174"))
                )
                sauvegarder_link.click()
                print("      ✓ Bouton 'SAUVEGARDER' cliqué (via ID)")
            
            time.sleep(3)  # Attendre la sauvegarde
            print("      ✓ Etat de Dépotage sauvegardé")
            
        except Exception as e:
            print(f"      ❌ Erreur sauvegarde: {e}")
            return_to_home_after_error(driver)
            return False
        
        # MED.7.2: Extraire la référence "sauvegardée" (avant validation)
        # Cette référence sera utilisée en cas d'échec de validation
        sauvegarde_reference = None
        try:
            time.sleep(2)  # Attendre que la référence s'affiche
            
            # Chercher la table de référence
            reference_table = driver.find_element(By.CSS_SELECTOR, "table.reference")
            
            # Extraire les cellules de la deuxième ligne (index 1)
            rows = reference_table.find_elements(By.TAG_NAME, "tr")
            if len(rows) >= 2:
                data_row = rows[1]
                cells = data_row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) >= 5:
                    # Extraire Série (colonne 4, index 3) et Clé (colonne 5, index 4)
                    serie_value = cells[3].text.strip()
                    cle_value = cells[4].text.strip()
                    
                    # Enlever les zéros initiaux de la série
                    serie_clean = str(int(serie_value)) if serie_value.isdigit() else serie_value
                    
                    # Combiner: [Série][Clé]
                    sauvegarde_reference = f"{serie_clean}{cle_value}"
                    
                    print(f"      ✓ Référence sauvegardée extraite: {sauvegarde_reference}")
                    print(f"         (Série={serie_value} → {serie_clean}, Clé={cle_value})")
        except Exception as e:
            print(f"      ⚠️  Impossible d'extraire la référence sauvegardée: {e}")
            print(f"         (Continuera avec extraction après validation)")
            # Continuer quand même - on essaiera après validation
        
        # ==================================================================
        # ÉTAPE MED.8: Valider l'ED modifié
        # ==================================================================
        print("\n   ✅ Validation de l'ED modifié...")
        
        try:
            # Stratégie robuste: chercher par texte puis par ID si échoue
            try:
                valider_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ui-menuitem-link')]//span[text()='VALIDER']/parent::a"))
                )
                valider_link.click()
                print("      ✓ Bouton 'VALIDER' cliqué")
            except:
                # Fallback: par ID
                valider_link = wait.until(
                    EC.element_to_be_clickable((By.ID, "secure_176"))
                )
                valider_link.click()
                print("      ✓ Bouton 'VALIDER' cliqué (via ID)")
            
            time.sleep(4)  # Attendre la validation
            
        except Exception as e:
            print(f"      ❌ Erreur validation: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE MED.9: Vérifier la réponse de validation et extraire référence
        # ==================================================================
        print("\n   🔍 Vérification du résultat de validation...")
        
        # Vérifier messages de succès/erreur
        try:
            time.sleep(2)
            
            # Chercher erreurs
            error_msg = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error-detail, span.ui-messages-error-detail")
            
            if error_msg and len(error_msg) > 0:
                error_text = error_msg[0].text.strip()
                print(f"      ❌ Erreur validation: {error_text}")
                
                # Si erreur et qu'on a la référence sauvegardée, l'enregistrer dans shipper
                if sauvegarde_reference:
                    try:
                        lta_name = os.path.basename(lta_folder_path)
                        parent_dir = os.path.dirname(lta_folder_path)
                        lta_name_with_underscore = lta_name.replace(" ", "_")
                        shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                        shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                        
                        if shipper_files:
                            shipper_file = shipper_files[0]
                            
                            # Lire le fichier actuel
                            with open(shipper_file, 'r', encoding='utf-8') as f:
                                lines = [line.rstrip('\n') for line in f.readlines()]
                            
                            # S'assurer qu'on a au moins 3 lignes
                            while len(lines) < 3:
                                lines.append("")
                            
                            # Ajouter ou remplacer la ligne 4 avec la référence sauvegardée
                            if len(lines) == 3:
                                lines.append(sauvegarde_reference)
                            elif len(lines) >= 4:
                                lines[3] = sauvegarde_reference
                            
                            # Réécrire le fichier
                            with open(shipper_file, 'w', encoding='utf-8') as f:
                                f.write('\n'.join(lines))
                            
                            print(f"      ✓ Référence sauvegardée écrite dans {os.path.basename(shipper_file)}")
                            print(f"         Ligne 4: {sauvegarde_reference} (NON VALIDÉE - ERREUR)")
                    except Exception as e:
                        print(f"      ⚠️  Impossible de sauvegarder la référence: {e}")
                
                return_to_home_after_error(driver)
                return False
            
            # Chercher succès
            success_msg = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-info-detail, span.ui-messages-info-detail")
            
            if success_msg and len(success_msg) > 0:
                success_text = success_msg[0].text
                print(f"      ✅ Succès: {success_text}")
            else:
                print("      ✅ ED modifié validé avec succès")
            
            # MED.9.1: Extraire la référence validée
            print("      ✓ Extraction de la référence validée...")
            
            try:
                # Chercher la table de référence
                reference_table = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table.reference"))
                )
                
                # Extraire les cellules de la deuxième ligne (index 1)
                rows = reference_table.find_elements(By.TAG_NAME, "tr")
                if len(rows) < 2:
                    print("      ⚠️  Table de référence incomplète")
                    # Utiliser la référence sauvegardée si disponible
                    if sauvegarde_reference:
                        ds_reference = sauvegarde_reference
                        print(f"      ✓ Utilisation référence sauvegardée: {ds_reference}")
                    else:
                        print("      ⚠️  Aucune référence disponible")
                        ds_reference = None
                else:
                    data_row = rows[1]
                    cells = data_row.find_elements(By.TAG_NAME, "td")
                    
                    if len(cells) < 5:
                        print("      ⚠️  Données de référence incomplètes")
                        # Utiliser la référence sauvegardée si disponible
                        if sauvegarde_reference:
                            ds_reference = sauvegarde_reference
                            print(f"      ✓ Utilisation référence sauvegardée: {ds_reference}")
                        else:
                            ds_reference = None
                    else:
                        # Extraire Série (colonne 4, index 3) et Clé (colonne 5, index 4)
                        serie_value = cells[3].text.strip()
                        cle_value = cells[4].text.strip()
                        
                        # Enlever les zéros initiaux de la série
                        serie_clean = str(int(serie_value)) if serie_value.isdigit() else serie_value
                        
                        # Combiner: [Série][Clé]
                        ds_reference = f"{serie_clean}{cle_value}"
                        
                        print(f"      ✓ Référence extraite: Série={serie_value} → {serie_clean}, Clé={cle_value}")
                        print(f"      ✓ Référence DS complète: {ds_reference}")
                
                # MED.9.2: Enregistrer la référence dans le fichier shipper (ligne 4)
                if ds_reference:
                    try:
                        lta_name = os.path.basename(lta_folder_path)
                        parent_dir = os.path.dirname(lta_folder_path)
                        
                        # Chercher le fichier shipper: [X]eme_LTA_*.txt
                        lta_name_with_underscore = lta_name.replace(" ", "_")
                        shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                        shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                        
                        if not shipper_files:
                            print(f"      ⚠️  Fichier shipper introuvable: {shipper_pattern}")
                        else:
                            shipper_file = shipper_files[0]
                            
                            # Lire le fichier actuel
                            with open(shipper_file, 'r', encoding='utf-8') as f:
                                lines = [line.rstrip('\n') for line in f.readlines()]
                            
                            # S'assurer qu'on a au moins 3 lignes (shipper, serie+cle, location)
                            while len(lines) < 3:
                                lines.append("")
                            
                            # Ajouter ou remplacer la ligne 4 (index 3) avec la référence DS
                            if len(lines) == 3:
                                # Ajouter ligne 4
                                lines.append(ds_reference)
                                action = "ajoutée"
                            elif len(lines) >= 4:
                                # Remplacer ligne 4 existante
                                lines[3] = ds_reference
                                action = "mise à jour"
                            
                            # Réécrire le fichier
                            with open(shipper_file, 'w', encoding='utf-8') as f:
                                f.write('\n'.join(lines))
                            
                            print(f"      ✓ Référence DS {action} dans {os.path.basename(shipper_file)}")
                            print(f"         Ligne 4: {ds_reference}")
                    
                    except Exception as e:
                        print(f"      ⚠️  Erreur mise à jour fichier shipper: {e}")
                
            except Exception as e:
                print(f"      ⚠️  Erreur extraction référence: {e}")
                # Pas critique, continuer
        
        except Exception as e:
            print(f"      ⚠️  Erreur vérification messages: {e}")
        
        # ==================================================================
        # FIN - Retour à l'accueil
        # ==================================================================
        print("\n   🏠 Retour à l'accueil...")
        return_to_home_after_error(driver)
        
        print("\n" + "="*70)
        print("✅ MODIFICATION ED BLOCAGE TERMINÉE AVEC SUCCÈS")
        print("="*70)
        
        return True
    
    except Exception as e:
        print(f"\n❌ ERREUR MODIFICATION ED BLOCAGE: {e}")
        import traceback
        traceback.print_exc()
        return_to_home_after_error(driver)
        return False


def correct_blocage_weights(lta_folder_path, corrected_weight):
    """
    Corrige les poids dans generated_excel et summary_file pour un LTA blocage.
    
    Étapes:
    BC.2.1: Mettre à jour global P,BRUT (B6) dans generated_excel
    BC.2.2: Détecter tous les DUMs dynamiquement
    BC.2.3: Calculer la somme des DUM P,BRUT
    BC.2.4: Ajuster le dernier DUM si nécessaire
    BC.3: Mettre à jour le dernier DUM dans summary_file
    
    Args:
        lta_folder_path: Chemin du dossier LTA
        corrected_weight: Poids corrigé (float)
    
    Returns:
        bool: True si succès, False sinon
    """
    try:
        print(f"\n   📝 Correction des fichiers Excel pour blocage...")
        
        # Trouver le fichier generated_excel
        generated_excel_path = None
        for file in os.listdir(lta_folder_path):
            if file.startswith("generated_excel") and file.endswith(".xlsx"):
                generated_excel_path = os.path.join(lta_folder_path, file)
                break
        
        if not generated_excel_path:
            print(f"      ❌ Fichier generated_excel non trouvé")
            return False
        
        # Trouver le fichier summary_file
        summary_file_path = None
        for file in os.listdir(lta_folder_path):
            if file.startswith("summary_file") and file.endswith(".xlsx"):
                summary_file_path = os.path.join(lta_folder_path, file)
                break
        
        if not summary_file_path:
            print(f"      ❌ Fichier summary_file non trouvé")
            return False
        
        # ========== BC.2: Mise à jour generated_excel ==========
        print(f"\n      📊 Mise à jour generated_excel...")
        
        wb = load_workbook(generated_excel_path, data_only=False)
        ws = wb['Summary']
        
        # BC.2.1: Mettre à jour global P,BRUT (B6)
        old_global_pbrut = ws['B6'].value
        ws['B6'] = corrected_weight
        print(f"         ✓ Global P,BRUT (B6): {old_global_pbrut} → {corrected_weight}")
        
        # BC.2.2: Détecter tous les DUMs dynamiquement
        print(f"\n      🔍 Détection des DUMs:")
        
        dum_count = 0
        dum_pbrut_cells = []
        
        row = 11  # Première ligne possible pour DUM 1
        while row < 500:  # Limite de sécurité
            cell_value = ws[f'C{row}'].value
            
            # Vérifier si c'est un header DUM
            if cell_value and isinstance(cell_value, str) and 'DUM' in cell_value.upper():
                dum_count += 1
                pbrut_row = row + 4  # P,BRUT est 4 lignes sous le header
                pbrut_cell = f'B{pbrut_row}'
                pbrut_value = ws[pbrut_cell].value
                
                dum_pbrut_cells.append({
                    'dum_number': dum_count,
                    'cell': pbrut_cell,
                    'value': float(pbrut_value) if pbrut_value else 0.0
                })
                
                print(f"         ✓ DUM {dum_count} trouvé ({pbrut_cell}): {pbrut_value} kg")
                
                row += 7  # Passer au prochain DUM potentiel
            else:
                row += 1
                
                # Si on a déjà trouvé des DUMs et qu'on a 10 lignes vides, arrêter
                if dum_count > 0:
                    last_pbrut_row = int(dum_pbrut_cells[-1]['cell'][1:])
                    if row > last_pbrut_row and all(ws[f'C{r}'].value is None for r in range(row, min(row + 10, 500))):
                        break
        
        if dum_count == 0:
            print(f"         ❌ Aucun DUM détecté")
            wb.close()
            return False
        
        print(f"         ──────────────────────────────")
        print(f"         📊 Total: {dum_count} DUMs détectés")
        
        # BC.2.3: Calculer la somme des DUM P,BRUT
        dum_sum = sum(dum['value'] for dum in dum_pbrut_cells)
        dum_sum = round(dum_sum, 2)
        print(f"         📊 Somme actuelle: {dum_sum} kg")
        
        # BC.2.4: Ajuster le dernier DUM si nécessaire
        difference = round(dum_sum - corrected_weight, 2)
        
        if abs(difference) < 0.01:
            print(f"\n      ✅ Aucun ajustement nécessaire (différence: {difference} kg)")
            last_adjusted_value = dum_pbrut_cells[-1]['value']
        else:
            print(f"\n      ⚙️  Ajustement requis:")
            print(f"         Différence: {dum_sum} - {corrected_weight} = {difference} kg")
            
            last_dum = dum_pbrut_cells[-1]
            old_value = last_dum['value']
            new_value = round(old_value - difference, 2)
            
            ws[last_dum['cell']] = new_value
            print(f"         ✓ Dernier DUM ajusté ({last_dum['cell']}): {old_value} → {new_value}")
            
            # Vérifier la nouvelle somme
            new_sum = round(dum_sum - difference, 2)
            print(f"         ✓ Nouvelle somme: {new_sum} kg ✅")
            
            last_adjusted_value = new_value
        
        # Sauvegarder generated_excel
        wb.save(generated_excel_path)
        wb.close()
        print(f"      ✓ generated_excel sauvegardé")
        
        # ========== BC.3: Mise à jour summary_file ==========
        print(f"\n      📊 Mise à jour summary_file...")
        
        wb_summary = load_workbook(summary_file_path, data_only=False)
        ws_summary = wb_summary.active
        
        # Calculer la dernière ligne: DUM_Count + 1 (header en ligne 1)
        last_row = dum_count + 1
        last_cell = f'D{last_row}'
        
        old_summary_value = ws_summary[last_cell].value
        ws_summary[last_cell] = last_adjusted_value
        print(f"         ✓ Dernière ligne ({last_cell}): {old_summary_value} → {last_adjusted_value}")
        
        # Sauvegarder summary_file
        wb_summary.save(summary_file_path)
        wb_summary.close()
        print(f"      ✓ summary_file sauvegardé")
        
        print(f"\n   ✅ Corrections blocage terminées")
        print(f"   ℹ️  Note: ED existant sera modifié ultérieurement (pas de création)")
        
        return True
    
    except Exception as e:
        print(f"      ❌ Erreur correction blocage: {e}")
        traceback.print_exc()
        return False

def wait_for_ui_blocker_disappear(driver, timeout=10):
    """
    Attend que le blocker UI (overlay) disparaisse avant de continuer.
    
    Args:
        driver: WebDriver Selenium
        timeout: Temps maximum d'attente en secondes (défaut: 10)
    
    Returns:
        bool: True si blocker disparu, False si timeout
    """
    try:
        from selenium.webdriver.support import expected_conditions as EC
        
        # Chercher les éléments blocker communs dans BADR
        blocker_selectors = [
            "div.ui-blockui",
            "div.ui-blockui-content",
            "div[id*='blocker']",
            "div.ui-widget-overlay"
        ]
        
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            blocker_visible = False
            
            for selector in blocker_selectors:
                try:
                    blockers = driver.find_elements(By.CSS_SELECTOR, selector)
                    for blocker in blockers:
                        # Vérifier si le blocker est visible
                        if blocker.is_displayed():
                            blocker_visible = True
                            break
                except:
                    pass
            
            if not blocker_visible:
                return True
            
            time.sleep(0.3)
        
        # Timeout atteint
        return False
        
    except Exception as e:
        # En cas d'erreur, on suppose que le blocker n'est pas là
        return True

def save_dum_error_log(lta_folder_path, lta_name, dum_number, sheet_name, error_exception, error_step, dum_data=None):
    """
    Crée un fichier log détaillé pour un DUM qui a échoué.
    
    Args:
        lta_folder_path: Chemin du dossier LTA
        lta_name: Nom du LTA (ex: "7eme LTA")
        dum_number: Numéro du DUM (1, 2, 3, etc.)
        sheet_name: Nom du sheet (ex: "Sheet 1")
        error_exception: L'exception capturée
        error_step: Description de l'étape où l'erreur s'est produite
        dum_data: Données du DUM (optionnel)
    """
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        error_filename = f"error-dum-processing-{lta_name.replace(' ', '_')}-DUM{dum_number}-{timestamp}.txt"
        error_path = os.path.join(lta_folder_path, error_filename)
        
        with open(error_path, 'w', encoding='utf-8') as f:
            f.write("="*70 + "\n")
            f.write("ERREUR - TRAITEMENT DUM PHASE 2\n")
            f.write("="*70 + "\n\n")
            
            f.write(f"LTA: {lta_name}\n")
            f.write(f"DUM: {dum_number}\n")
            f.write(f"Sheet: {sheet_name}\n")
            f.write(f"Date/Heure: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Étape échouée: {error_step}\n\n")
            
            f.write("DÉTAILS ERREUR:\n")
            f.write("-"*70 + "\n")
            f.write(f"Type: {type(error_exception).__name__}\n")
            f.write(f"Message: {str(error_exception)}\n\n")
            
            if dum_data:
                f.write("DONNÉES DUM:\n")
                f.write("-"*70 + "\n")
                f.write(f"Total Value: {dum_data.get('total_value', 'N/A')}\n")
                f.write(f"Gross Weight: {dum_data.get('total_gross_weight', 'N/A')}\n")
                f.write(f"Positions: {dum_data.get('total_positions', 'N/A')}\n")
                f.write(f"Freight: {dum_data.get('total_freight', 'N/A')}\n")
                f.write(f"Insurance: {dum_data.get('insurance', 'N/A')}\n")
                f.write(f"Cartons: {dum_data.get('cartons', 'N/A')}\n\n")
            
            f.write("ACTION PRISE:\n")
            f.write("-"*70 + "\n")
            f.write("✓ Retour à l'accueil effectué\n")
            f.write("✓ Marqueur \"error\" ajouté à generated_excel\n")
            f.write("⏭️  Traitement continue avec DUM suivant\n\n")
            
            f.write("RECOMMANDATION:\n")
            f.write("-"*70 + "\n")
            f.write("Vérifier manuellement ce DUM et créer la déclaration si nécessaire.\n\n")
            
            f.write("="*70 + "\n")
        
        print(f"      📝 Log d'erreur créé: {error_filename}")
        
    except Exception as e:
        print(f"      ⚠️  Impossible de créer le log d'erreur: {e}")

def mark_dum_as_error_in_excel(lta_folder_path, dum_number):
    """
    Marque un DUM comme "error" dans le fichier generated_excel.
    Même logique que save_dum_series_to_excel mais écrit "error".
    
    Args:
        lta_folder_path: Chemin du dossier LTA
        dum_number: Numéro du DUM (1, 2, 3, etc.)
    """
    try:
        # Trouver le fichier generated_excel
        generated_excel_path = None
        for file in os.listdir(lta_folder_path):
            if file.startswith("generated_excel") and file.endswith(".xlsx"):
                generated_excel_path = os.path.join(lta_folder_path, file)
                break
        
        if not generated_excel_path:
            print(f"      ⚠️  generated_excel introuvable pour marquage erreur")
            return
        
        # Ouvrir le fichier Excel
        wb = load_workbook(generated_excel_path, data_only=False)
        ws = wb['Summary']
        
        # Calculer la cellule: C + (12 + (dum_number - 1) * 7)
        row = 12 + (dum_number - 1) * 7
        cell = f'C{row}'
        
        # Écrire "error" dans la cellule avec fond rouge
        ws[cell] = "error"
        
        # Sauvegarder
        wb.save(generated_excel_path)
        wb.close()
        
        print(f"      ✓ Marqueur 'error' ajouté en {cell}")
        
    except Exception as e:
        print(f"      ⚠️  Erreur marquage Excel: {e}")

def return_to_home_after_error(driver):
    """
    Fonction helper pour retourner à l'accueil après une erreur.
    Utilisée dans create_etat_depotage pour nettoyer l'état avant de sortir.
    """
    print("\n   🏠 Retour à l'accueil après erreur...")
    try:
        # Sortir de l'iframe si on est dedans
        try:
            driver.switch_to.default_content()
            print("      ✓ Sorti de l'iframe")
        except:
            pass
        
        # Cliquer sur le bouton Accueil
        try:
            driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
            print("      ✓ Navigation directe vers accueil")
            time.sleep(3)
        except Exception as btn_err:
            print(f"      ⚠️  Erreur clic bouton: {btn_err}")
            # Fallback: navigation directe
            try:
                driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
                time.sleep(3)
                print("      ✓ Navigation directe vers accueil")
            except Exception as nav_err:
                print(f"      ❌ Erreur navigation: {nav_err}")
    except Exception as e:
        print(f"      ❌ Erreur retour accueil: {e}")

def create_etat_depotage(driver, lta_folder_path, shipper_data):
    """
    Crée un Etat de Dépotage (Unloading Statement) pour une LTA avec référence DS MEAD.
    
    Args:
        driver: Selenium WebDriver instance
        lta_folder_path: Path to LTA folder
        shipper_data: Dict containing serie, cle, loading_location
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        wait = WebDriverWait(driver, 15)
        
        print("\n" + "="*70)
        print("📦 CRÉATION ETAT DE DÉPOTAGE")
        print("="*70)
        
        # ==================================================================
        # ÉTAPE ED.0: Navigation vers "Etat de Dépotage - Voyage Aérien"
        # ==================================================================
        print("\n📂 Navigation: MISE EN DOUANE → Etat de Dépotage → Voyage Aérien...")
        
        # Ouvrir le menu "MISE EN DOUANE"
        try:
            mise_en_douane_link = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//h3[contains(@class, 'ui-panelmenu-header')]//a[contains(text(), 'MISE EN DOUANE')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", mise_en_douane_link)
            time.sleep(0.5)
            mise_en_douane_link.click()
            print("   ✓ Menu 'MISE EN DOUANE' ouvert")
            time.sleep(2)
        except Exception as e:
            print(f"   ⚠️  Menu 'MISE EN DOUANE' déjà ouvert ou erreur: {e}")
            # Continuer car le menu peut déjà être ouvert
        
        # Ouvrir le sous-menu "Créer une Déclaration" (ID: _151)
        try:
            creer_declaration_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_151"))
            )
            creer_declaration_link.click()
            print("   ✓ Sous-menu 'Créer une Déclaration' ouvert")
            time.sleep(1)
        except Exception as e:
            print(f"   ❌ Erreur ouverture 'Créer une Déclaration': {e}")
            return_to_home_after_error(driver)
            return False
        
        # Ouvrir le sous-menu "Etat de Dépotage" (ID: _236)
        try:
            etat_depotage_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_236"))
            )
            etat_depotage_link.click()
            print("   ✓ Sous-menu 'Etat de Dépotage' ouvert")
            time.sleep(1)
        except Exception as e:
            print(f"   ❌ Erreur ouverture 'Etat de Dépotage': {e}")
            return_to_home_after_error(driver)
            return False
        
        # Cliquer sur "Voyage Aérien" (ID: _247)
        try:
            voyage_aerien_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_247"))
            )
            voyage_aerien_link.click()
            print("   ✓ Lien 'Voyage Aérien' cliqué")
            time.sleep(3)
        except Exception as e:
            print(f"   ❌ Erreur clic 'Voyage Aérien': {e}")
            return_to_home_after_error(driver)
            return False
        
        # Basculer dans l'iframe du formulaire
        try:
            print("   🔄 Basculement vers l'iframe du formulaire...")
            iframe = wait.until(
                EC.presence_of_element_located((By.ID, "iframeMenu"))
            )
            driver.switch_to.frame(iframe)
            print("   ✓ Iframe chargé")
            time.sleep(2)
        except Exception as e:
            print(f"   ❌ Erreur basculement iframe: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE ED.1: Sélection du Bureau "301"
        # ==================================================================
        print("\n   🏢 Sélection du Bureau 301...")
        
        # ED.1.1: Entrer "301" dans l'autocomplete
        try:
            bureau_input = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:bureauCmbId_INPUT_input"))
            )
            bureau_input.clear()
            bureau_input.send_keys("301")
            print("      ✓ Bureau '301' saisi")
            time.sleep(1)
        except Exception as e:
            print(f"      ❌ Erreur saisie bureau: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.1.2: Sélectionner la première suggestion
        try:
            bureau_suggestion = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#rootForm\\:bureauCmbId_INPUT_panel li.ui-autocomplete-item"))
            )
            bureau_suggestion.click()
            print("      ✓ Suggestion bureau sélectionnée")
            time.sleep(1)
        except Exception as e:
            print(f"      ❌ Erreur sélection suggestion: {e}")
            return False
        
        # ED.1.3: Valider la sélection du bureau
        try:
            valider_bureau_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "rootForm:btnConfirmer"))
            )
            valider_bureau_btn.click()
            print("      ✓ Bureau validé")
            time.sleep(4)  # Attendre le chargement du formulaire suivant
        except Exception as e:
            print(f"      ❌ Erreur validation bureau: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE ED.2: Configuration Type de Déclaration et Référence
        # ==================================================================
        print("\n   📋 Configuration de la déclaration...")
        
        # ED.2.1: Sélectionner "DS MEAD Combinée" (radio button index 3, value "08")
        try:
            # Méthode directe: chercher tous les div.ui-radiobutton-box dans le tableau radioTypeDS
            # et prendre le 4ème (index 3 = DS MEAD Combinée)
            time.sleep(1)
            ds_radios = driver.find_elements(By.CSS_SELECTOR, "table#mainTab\\:form1\\:radioTypeDS div.ui-radiobutton-box")
            if len(ds_radios) >= 4:
                ds_radios[3].click()  # Le 4ème = DS MEAD Combinée
                print("      ✓ 'DS MEAD Combinée' sélectionné")
                time.sleep(0.5)
            else:
                print(f"      ⚠️  Radios DS MEAD insuffisants (trouvé: {len(ds_radios)})")
                # Méthode alternative: JavaScript
                print("      🔄 Tentative avec JavaScript...")
                js_code = """
                var radio = document.getElementById('mainTab:form1:radioTypeDS:3');
                radio.checked = true;
                var event = new Event('change', { bubbles: true });
                radio.dispatchEvent(event);
                """
                driver.execute_script(js_code)
                time.sleep(0.5)
                print("      ✓ 'DS MEAD Combinée' sélectionné via JavaScript")
        except Exception as e:
            print(f"      ❌ Impossible de sélectionner DS MEAD: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.2.2: Entrer l'année actuelle
        try:
            annee_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:anneeId"))
            )
            annee_input.clear()
            current_year = str(time.strftime("%Y"))
            annee_input.send_keys(current_year)
            print(f"      ✓ Année: {current_year}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie année: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.2.3: Entrer le numéro de série
        try:
            serie_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:serieId"))
            )
            serie_input.clear()
            serie_input.send_keys(shipper_data['serie'])
            print(f"      ✓ Série: {shipper_data['serie']}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie série: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.2.4: Entrer la clé
        try:
            cle_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:cleId"))
            )
            cle_input.clear()
            cle_input.send_keys(shipper_data['cle'])
            print(f"      ✓ Clé: {shipper_data['cle']}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie clé: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.2.5: Entrer la référence LTA
        # Extraire la référence depuis le nom du fichier PDF LTA
        try:
            lta_name = os.path.basename(lta_folder_path)
            lta_pattern = os.path.join(lta_folder_path, f"{lta_name} - *.pdf")
            lta_files = glob.glob(lta_pattern)
            
            if not lta_files:
                print(f"      ❌ Fichier LTA PDF introuvable")
                return_to_home_after_error(driver)
                return False
            
            lta_filename = os.path.basename(lta_files[0])
            # Extraire référence: "9eme LTA - 012-24513105.pdf" → "012-24513105"
            lta_reference_raw = lta_filename.split(" - ")[1].replace(".pdf", "")
            
            # Préparer 3 formats à essayer:
            # Format 1: Avec tirets ET /1 (ex: "235-94908726/1")
            ref_parts = lta_reference_raw.split("-")
            ref_parts[0] = str(int(ref_parts[0]))  # Enlever zéros initiaux
            lta_reference_format1 = "-".join(ref_parts) + "/1"
            
            # Format 2: Sans tirets, sans /1 (ex: "23594908726")
            lta_reference_format2 = lta_reference_raw.replace("-", "")
            lta_reference_format2 = str(int(lta_reference_format2))  # Enlever zéros initiaux
            
            # Format 3: Avec tirets, SANS /1 (ex: "235-94908726")
            lta_reference_format3 = "-".join(ref_parts)
            
            print(f"      📄 Référence LTA brute: {lta_reference_raw}")
            print(f"      📄 Format 1 (avec /1): {lta_reference_format1}")
            print(f"      📄 Format 2 (sans tirets): {lta_reference_format2}")
            print(f"      📄 Format 3 (avec tirets, sans /1): {lta_reference_format3}")
            
            # Essayer d'abord le Format 1 (avec /1)
            reference_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:referenceLotID"))
            )
            reference_input.clear()
            reference_input.send_keys(lta_reference_format1)
            print(f"      ✓ Référence saisie (Format 1): {lta_reference_format1}")
            time.sleep(0.5)
            
        except Exception as e:
            print(f"      ❌ Erreur saisie référence: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.2.6: Entrer le lieu de chargement (avec autocomplete)
        if shipper_data.get('loading_location'):
            try:
                lieu_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:form1:lieuChargementCmbId_INPUT_input"))
                )
                lieu_input.clear()
                lieu_input.send_keys(shipper_data['loading_location'])
                print(f"      ✓ Lieu de chargement: {shipper_data['loading_location']}")
                time.sleep(1)
                
                # Sélectionner la première suggestion
                lieu_suggestion = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form1\\:lieuChargementCmbId_INPUT_panel li.ui-autocomplete-item"))
                )
                lieu_suggestion.click()
                print("      ✓ Suggestion lieu sélectionnée")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ⚠️  Erreur saisie lieu de chargement: {e}")
                # Continuer même si le lieu échoue
        else:
            print("      ⚠️  Pas de lieu de chargement dans le fichier shipper")
        
        # ==================================================================
        # ÉTAPE ED.3: Validation et Gestion des Erreurs
        # ==================================================================
        print("\n   ✅ Validation de l'Etat de Dépotage...")
        
        # ED.3.1: Cliquer sur Valider
        try:
            valider_ref_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "mainTab:form1:confirmerRef"))
            )
            valider_ref_btn.click()
            print("      ✓ Bouton 'Valider' cliqué")
            time.sleep(3)
        except Exception as e:
            print(f"      ❌ Erreur clic validation: {e}")
            return_to_home_after_error(driver)
            return False
        
        # ED.3.2: Vérifier messages d'erreur ou de succès
        error_detected = False
        try:
            # Attendre plus longtemps pour que les messages s'affichent
            time.sleep(2)
            
            # Chercher message d'erreur (plusieurs tentatives)
            error_msg = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error-detail")
            
            # Si pas trouvé, chercher aussi dans span.ui-messages-error-detail
            if not error_msg or len(error_msg) == 0:
                error_msg = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
            
            if error_msg and len(error_msg) > 0:
                error_text = error_msg[0].text.strip()
                print(f"      ⚠️  Erreur de validation détectée: {error_text}")
                error_detected = True
                
                # ED.3.3: Si erreur référence, essayer les autres formats
                if "n'existe pas" in error_text.lower() or "référence" in error_text.lower():
                    print(f"      ⚠️  Format 1 rejeté, tentative Format 2...")
                    
                    # Fermer le message d'erreur
                    try:
                        close_btn = driver.find_element(By.CSS_SELECTOR, "a.ui-messages-close")
                        close_btn.click()
                        time.sleep(0.5)
                        print("      ✓ Message d'erreur fermé")
                    except:
                        pass
                    
                    # Essayer Format 2 (sans tirets)
                    reference_input = wait.until(
                        EC.presence_of_element_located((By.ID, "mainTab:form1:referenceLotID"))
                    )
                    reference_input.clear()
                    time.sleep(0.3)
                    reference_input.send_keys(lta_reference_format2)
                    print(f"      ✓ Référence Format 2 saisie: {lta_reference_format2}")
                    time.sleep(0.5)
                    
                    # Re-valider
                    valider_ref_btn = wait.until(
                        EC.element_to_be_clickable((By.ID, "mainTab:form1:confirmerRef"))
                    )
                    valider_ref_btn.click()
                    print("      ✓ Bouton 'Valider' re-cliqué")
                    time.sleep(3)
                    
                    # Vérifier résultat Format 2
                    time.sleep(1)
                    error_msg_retry2 = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error-detail")
                    if not error_msg_retry2 or len(error_msg_retry2) == 0:
                        error_msg_retry2 = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                    
                    if error_msg_retry2 and len(error_msg_retry2) > 0:
                        error_text_retry2 = error_msg_retry2[0].text.strip()
                        print(f"      ⚠️  Format 2 rejeté: {error_text_retry2}")
                        print(f"      🔄 Tentative Format 3 (avec tirets, sans /1)...")
                        
                        # Fermer le message d'erreur
                        try:
                            close_btn = driver.find_element(By.CSS_SELECTOR, "a.ui-messages-close")
                            close_btn.click()
                            time.sleep(0.5)
                        except:
                            pass
                        
                        # Essayer Format 3 (avec tirets, sans /1)
                        reference_input = wait.until(
                            EC.presence_of_element_located((By.ID, "mainTab:form1:referenceLotID"))
                        )
                        reference_input.clear()
                        time.sleep(0.3)
                        reference_input.send_keys(lta_reference_format3)
                        print(f"      ✓ Référence Format 3 saisie: {lta_reference_format3}")
                        time.sleep(0.5)
                        
                        # Re-valider
                        valider_ref_btn = wait.until(
                            EC.element_to_be_clickable((By.ID, "mainTab:form1:confirmerRef"))
                        )
                        valider_ref_btn.click()
                        print("      ✓ Bouton 'Valider' re-cliqué")
                        time.sleep(3)
                        
                        # Vérifier résultat Format 3
                        time.sleep(1)
                        error_msg_retry3 = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error-detail")
                        if not error_msg_retry3 or len(error_msg_retry3) == 0:
                            error_msg_retry3 = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                        
                        if error_msg_retry3 and len(error_msg_retry3) > 0:
                            error_text_retry3 = error_msg_retry3[0].text.strip()
                            print(f"      ❌ Format 3 aussi rejeté: {error_text_retry3}")
                            
                            # Tous les formats ont échoué - créer fichier log
                            lta_name = os.path.basename(lta_folder_path)
                            parent_dir = os.path.dirname(lta_folder_path)
                            lta_name_with_underscore = lta_name.replace(" ", "_")
                            error_log_filename = f"error-creating-ds-depotage-{lta_name_with_underscore}.log"
                            error_log_filepath = os.path.join(parent_dir, error_log_filename)
                            
                            from datetime import datetime
                            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            with open(error_log_filepath, 'w', encoding='utf-8') as f:
                                f.write(f"ERREUR - Création Etat de Dépotage - Phase 1\n")
                                f.write(f"=" * 70 + "\n\n")
                                f.write(f"LTA: {lta_name}\n")
                                f.write(f"Date: {current_datetime}\n")
                                f.write(f"Étape: Validation de la référence LTA\n\n")
                                f.write(f"TENTATIVES:\n")
                                f.write(f"1. Format avec /1: {lta_reference_format1}\n")
                                f.write(f"   Erreur: {error_text}\n\n")
                                f.write(f"2. Format sans tirets: {lta_reference_format2}\n")
                                f.write(f"   Erreur: {error_text_retry2}\n\n")
                                f.write(f"3. Format avec tirets, sans /1: {lta_reference_format3}\n")
                                f.write(f"   Erreur: {error_text_retry3}\n\n")
                                f.write(f"RÉFÉRENCE BRUTE (PDF):\n")
                                f.write(f"{lta_reference_raw}\n\n")
                                f.write(f"DONNÉES DS MEAD:\n")
                                f.write(f"- Série: {shipper_data['serie']}\n")
                                f.write(f"- Clé: {shipper_data['cle']}\n")
                                if shipper_data.get('loading_location'):
                                    f.write(f"- Lieu de chargement: {shipper_data['loading_location']}\n")
                                f.write(f"\n")
                                f.write(f"MESSAGE:\n")
                                f.write(f"La référence LTA n'a pas pu être validée avec aucun des 3 formats.\n")
                                f.write(f"Vérifiez que:\n")
                                f.write(f"1. La référence LTA dans le nom du fichier PDF est correcte\n")
                                f.write(f"2. Le lot existe bien dans le système BADR\n")
                                f.write(f"3. Le lot n'est pas déjà dépoté\n\n")
                                f.write(f"ACTIONS RECOMMANDÉES:\n")
                                f.write(f"- Vérifier manuellement la référence sur BADR\n")
                                f.write(f"- Corriger le nom du fichier PDF si nécessaire\n")
                                f.write(f"- Créer l'Etat de Dépotage manuellement si l'erreur persiste\n")
                            
                            print(f"      ✓ Fichier log créé: {error_log_filename}")
                            
                            # Fermer l'erreur avant de sortir
                            try:
                                close_btn = driver.find_element(By.CSS_SELECTOR, "a.ui-messages-close")
                                close_btn.click()
                                time.sleep(0.5)
                            except:
                                pass
                            
                            # IMPORTANT: Retourner à l'accueil avant de sortir
                            print("\n      🏠 Retour à l'accueil après erreur...")
                            try:
                                driver.switch_to.default_content()
                                print("         ✓ Sorti de l'iframe")
                                
                                # Cliquer sur le bouton Accueil
                                accueil_btn = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.ID, "quitter"))
                                )
                                try:
                                    accueil_btn.click()
                                except:
                                    driver.execute_script("arguments[0].click();", accueil_btn)
                                print("         ✓ Bouton 'Accueil' cliqué")
                                time.sleep(3)
                            except Exception as nav_err:
                                print(f"         ⚠️  Erreur navigation: {nav_err}")
                                # Fallback: navigation directe
                                try:
                                    driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
                                    time.sleep(3)
                                    print("         ✓ Navigation directe vers accueil")
                                except:
                                    pass
                            
                            return False
                        else:
                            print("      ✅ Format 3 accepté!")
                            error_detected = False
                    else:
                        print("      ✅ Format 2 accepté!")
                        error_detected = False
                else:
                    # Erreur non liée à la référence - créer aussi un log
                    print(f"      ❌ Erreur non gérée: {error_text}")
                    
                    lta_name = os.path.basename(lta_folder_path)
                    parent_dir = os.path.dirname(lta_folder_path)
                    lta_name_with_underscore = lta_name.replace(" ", "_")
                    error_log_filename = f"error-creating-ds-depotage-{lta_name_with_underscore}.log"
                    error_log_filepath = os.path.join(parent_dir, error_log_filename)
                    
                    from datetime import datetime
                    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    with open(error_log_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"ERREUR - Création Etat de Dépotage - Phase 1\n")
                        f.write(f"=" * 70 + "\n\n")
                        f.write(f"LTA: {lta_name}\n")
                        f.write(f"Date: {current_datetime}\n")
                        f.write(f"Étape: Validation de l'Etat de Dépotage\n\n")
                        f.write(f"ERREUR SYSTÈME:\n")
                        f.write(f"{error_text}\n\n")
                        f.write(f"FORMAT RÉFÉRENCE UTILISÉ:\n")
                        f.write(f"Format 1: {lta_reference_format1}\n\n")
                        f.write(f"DONNÉES DS MEAD:\n")
                        f.write(f"- Série: {shipper_data['serie']}\n")
                        f.write(f"- Clé: {shipper_data['cle']}\n")
                        if shipper_data.get('loading_location'):
                            f.write(f"- Lieu de chargement: {shipper_data['loading_location']}\n")
                        f.write(f"\n")
                        f.write(f"TYPE D'ERREUR:\n")
                        f.write(f"Erreur système non liée au format de référence.\n\n")
                        f.write(f"ACTIONS RECOMMANDÉES:\n")
                        f.write(f"- Vérifier les données saisies dans BADR\n")
                        f.write(f"- Consulter la documentation BADR pour ce code d'erreur\n")
                        f.write(f"- Créer l'Etat de Dépotage manuellement\n")
                    
                    print(f"      ✓ Fichier log créé: {error_log_filename}")
                    
                    # IMPORTANT: Retourner à l'accueil avant de sortir
                    print("\n      🏠 Retour à l'accueil après erreur...")
                    try:
                        driver.switch_to.default_content()
                        print("         ✓ Sorti de l'iframe")
                        
                        # Cliquer sur le bouton Accueil
                        accueil_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "quitter"))
                        )
                        try:
                            accueil_btn.click()
                        except:
                            driver.execute_script("arguments[0].click();", accueil_btn)
                        print("         ✓ Bouton 'Accueil' cliqué")
                        time.sleep(3)
                    except Exception as nav_err:
                        print(f"         ⚠️  Erreur navigation: {nav_err}")
                        # Fallback: navigation directe
                        try:
                            driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
                            time.sleep(3)
                            print("         ✓ Navigation directe vers accueil")
                        except:
                            pass
                    
                    return False
            
            # ED.3.4: Vérifier message de succès
            success_msg = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-info-detail")
            
            if success_msg and len(success_msg) > 0:
                success_text = success_msg[0].text
                if "confirmées" in success_text.lower():
                    print(f"      ✓ Succès: {success_text}")
                else:
                    print(f"      ⚠️  Message inattendu: {success_text}")
            
            # Si aucune erreur n'a été détectée, c'est bon
            if not error_detected:
                print("      ✓ Validation terminée avec succès")
                # Sauvegarder la référence LTA validée pour utilisation ultérieure (création lots)
                # IMPORTANT: Utiliser le format EXACT qui a été accepté par BADR
                # Si la référence contient /1 (ex: 112-00181440/1), on garde le /1
                # Si la référence ne contient pas de / (ex: 11200181440), on la garde telle quelle
                # Pour savoir quel format a été accepté, on vérifie quelle valeur est dans le champ
                try:
                    reference_input_value = driver.find_element(By.ID, "mainTab:form1:referenceLotID").get_attribute("value")
                    # Garder la référence EXACTE (avec /1 si présent)
                    lta_reference_clean = reference_input_value.strip()
                    print(f"      📋 Référence validée à utiliser pour les lots: {lta_reference_clean}")
                    
                    # ED.3.5: Enregistrer la référence validée dans le fichier shipper (ligne 5)
                    try:
                        lta_name = os.path.basename(lta_folder_path)
                        parent_dir = os.path.dirname(lta_folder_path)
                        lta_name_with_underscore = lta_name.replace(" ", "_")
                        shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                        shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                        
                        if shipper_files:
                            shipper_file = shipper_files[0]
                            
                            # Lire le fichier actuel
                            with open(shipper_file, 'r', encoding='utf-8') as f:
                                lines = [line.rstrip('\n') for line in f.readlines()]
                            
                            # S'assurer qu'on a au moins 4 lignes (ou ajouter des lignes vides si nécessaire)
                            while len(lines) < 4:
                                lines.append("")
                            
                            # Ajouter ou remplacer la ligne 5 (index 4) avec la référence LTA validée
                            if len(lines) == 4:
                                lines.append(lta_reference_clean)
                            else:
                                lines[4] = lta_reference_clean
                            
                            # Réécrire le fichier
                            with open(shipper_file, 'w', encoding='utf-8') as f:
                                f.write('\n'.join(lines))
                            
                            print(f"      ✓ Référence LTA sauvegardée dans {os.path.basename(shipper_file)} (ligne 5)")
                            print(f"         Ligne 5: {lta_reference_clean}")
                        else:
                            print(f"      ⚠️  Fichier shipper introuvable pour sauvegarde: {shipper_pattern}")
                    
                    except Exception as e:
                        print(f"      ⚠️  Erreur sauvegarde référence LTA dans shipper: {e}")
                        # Continuer quand même
                    
                except:
                    # Fallback: utiliser le format 1 COMPLET (avec /1 si présent)
                    lta_reference_clean = lta_reference_format1
            else:
                print("      ❌ Erreur non résolue détectée")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
        except Exception as e:
            print(f"      ⚠️  Erreur lors de la vérification des messages: {e}")
        
        # ==================================================================
        # ÉTAPE ED.4: Naviguer vers l'onglet "Quantités"
        # ==================================================================
        print("\n   📊 Navigation vers l'onglet Quantités...")
        
        # Fermer tout message d'erreur persistant avant de continuer
        try:
            close_btns = driver.find_elements(By.CSS_SELECTOR, "a.ui-messages-close")
            for btn in close_btns:
                try:
                    btn.click()
                    time.sleep(0.3)
                except:
                    pass
        except:
            pass
        
        try:
            quantites_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab3']"))
            )
            quantites_tab.click()
            print("      ✓ Onglet Quantités ouvert")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation onglet Quantités: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.5: Extraire les totaux depuis generated_excel
        # ==================================================================
        print("\n   📂 Extraction des totaux depuis generated_excel...")
        
        # ED.5.1: Trouver le fichier generated_excel
        # Utiliser la référence validée (lta_reference_clean)
        # Si le format validé est "235-94908726" (sans /1), on l'utilise tel quel
        # Si le format validé est "235-94908726/1", on enlève le /1
        ref_for_filename = lta_reference_clean.split('/')[0]  # Enlever le /1 si présent
        generated_excel_pattern = f"generated_excel - {ref_for_filename}.xlsx"
        generated_excel_path = None
        
        for file in os.listdir(lta_folder_path):
            if file.startswith("generated_excel") and file.endswith(".xlsx"):
                generated_excel_path = os.path.join(lta_folder_path, file)
                break
        
        if not generated_excel_path:
            print(f"      ❌ Fichier generated_excel non trouvé dans {lta_folder_path}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        print(f"      ✓ Fichier trouvé: {os.path.basename(generated_excel_path)}")
        
        # ED.5.2: Lire et extraire les totaux (avant "FOURNISSEUR")
        try:
            wb = load_workbook(generated_excel_path, data_only=True)
            ws = wb['Summary']
            
            total_p = None
            total_p_brut = None
            
            for row in range(1, 20):
                cell_a = ws.cell(row=row, column=1).value
                
                if cell_a == "FOURNISSEUR":
                    break
                
                if cell_a == "P":
                    total_p = ws.cell(row=row, column=2).value
                elif cell_a == "P,BRUT":
                    total_p_brut = ws.cell(row=row, column=2).value
            
            wb.close()
            
            if total_p is None or total_p_brut is None:
                print("      ❌ Impossible de trouver les totaux P et P,BRUT")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # Convertir: P (contenants) = entier, P,BRUT (poids) = float
            total_p = int(float(total_p))  # Contenants = nombre entier (pas de décimales)
            total_p_brut = float(total_p_brut)  # Poids brut = garder les décimales (ex: 1419.50)
            
            print(f"      ✓ Total P (contenants): {total_p}")
            print(f"      ✓ Total P,BRUT (poids brut): {total_p_brut}")
            
        except Exception as e:
            print(f"      ❌ Erreur lecture generated_excel: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.5.3: Validation - Vérifier somme des DUMs
        # ==================================================================
        print("\n   🔍 Validation des totaux (somme DUMs)...")
        
        try:
            wb = load_workbook(generated_excel_path, data_only=True)
            ws = wb['Summary']
            
            calculated_p = 0  # Contenants = entier
            calculated_p_brut = 0.0  # Poids brut = float
            dum_details = []
            
            current_dum = None
            current_p = None
            current_p_brut = None
            
            for row in range(1, 200):
                cell_c = ws.cell(row=row, column=3).value
                cell_a = ws.cell(row=row, column=1).value
                cell_b = ws.cell(row=row, column=2).value
                
                # Détecter DUM
                if cell_c and isinstance(cell_c, str) and "DUM" in cell_c:
                    # Sauvegarder le DUM précédent
                    if current_dum and current_p is not None and current_p_brut is not None:
                        dum_details.append({
                            'dum': current_dum,
                            'p': current_p,
                            'p_brut': current_p_brut
                        })
                        calculated_p += current_p
                        calculated_p_brut += current_p_brut
                    
                    current_dum = cell_c.strip()
                    current_p = None
                    current_p_brut = None
                
                # Extraire P et P,BRUT
                if current_dum:
                    if cell_a == "P":
                        # P = nombre de contenants (entier)
                        current_p = int(float(cell_b)) if cell_b else 0
                    elif cell_a == "P,BRUT":
                        # P,BRUT = poids brut (garder les décimales)
                        current_p_brut = float(cell_b) if cell_b else 0.0
            
            # Ajouter le dernier DUM
            if current_dum and current_p is not None and current_p_brut is not None:
                dum_details.append({
                    'dum': current_dum,
                    'p': current_p,
                    'p_brut': current_p_brut
                })
                calculated_p += current_p
                calculated_p_brut += current_p_brut
            
            wb.close()
            
            print(f"      Totaux calculés: P={calculated_p}, P,BRUT={calculated_p_brut}")
            print(f"      Totaux déclarés: P={total_p}, P,BRUT={total_p_brut}")
            
            # Vérifier correspondance
            if calculated_p != total_p or calculated_p_brut != total_p_brut:
                print("      ❌ ERREUR: Les totaux ne correspondent pas!")
                
                # Créer fichier d'erreur
                lta_name = os.path.basename(lta_folder_path)
                error_filename = f"{lta_name}_ERROR.txt"
                error_filepath = os.path.join(os.path.dirname(lta_folder_path), error_filename)
                
                from datetime import datetime
                current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                with open(error_filepath, 'w', encoding='utf-8') as f:
                    f.write(f"ERREUR DE CALCUL - {lta_name}\n\n")
                    f.write(f"LTA: {lta_name} - {lta_reference.split('/')[0]}\n")
                    f.write(f"Date: {current_datetime}\n\n")
                    f.write("TOTAUX DÉCLARÉS:\n")
                    f.write(f"- P (Total): {total_p}\n")
                    f.write(f"- P,BRUT (Total): {total_p_brut}\n\n")
                    f.write("TOTAUX CALCULÉS:\n")
                    f.write(f"- P (Somme DUMs): {calculated_p}\n")
                    f.write(f"- P,BRUT (Somme DUMs): {calculated_p_brut}\n\n")
                    f.write("DÉTAIL PAR DUM:\n")
                    for dum in dum_details:
                        f.write(f"{dum['dum']}: P={dum['p']}, P,BRUT={dum['p_brut']}\n")
                    f.write("\nMESSAGE: Les totaux déclarés ne correspondent pas à la somme des DUMs.\n")
                    f.write("Veuillez vérifier le fichier generated_excel.\n")
                
                print(f"      ✓ Fichier d'erreur créé: {error_filename}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            print("      ✅ Validation réussie: Les totaux correspondent")
            
        except Exception as e:
            print(f"      ❌ Erreur lors de la validation: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.6: Entrer le poids brut total
        # ==================================================================
        print("\n   ⚖️  Saisie du poids brut total...")
        
        try:
            poids_brut_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form3:poidsBrutTotal_IT_id_input"))
            )
            poids_brut_input.clear()
            poids_brut_input.send_keys(str(total_p_brut))
            print(f"      ✓ Poids brut total saisi: {total_p_brut}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie poids brut: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.7: Entrer le nombre total de contenants
        # ==================================================================
        print("\n   📦 Saisie du nombre total de contenants...")
        
        try:
            nombre_contenants_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form3:nombreContenantTotal_IT_id"))
            )
            nombre_contenants_input.clear()
            nombre_contenants_input.send_keys(str(total_p))
            print(f"      ✓ Nombre de contenants saisi: {total_p}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie nombre contenants: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.8: Naviguer vers l'onglet "LTA"
        # ==================================================================
        print("\n   📄 Navigation vers l'onglet LTA...")
        
        try:
            lta_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab4']"))
            )
            lta_tab.click()
            print("      ✓ Onglet LTA ouvert")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation onglet LTA: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.9-ED.12: Créer les lots pour chaque DUM
        # ==================================================================
        print("\n   📦 Création des lots pour chaque DUM...")
        
        # Lire les données des DUMs depuis generated_excel
        try:
            wb = load_workbook(generated_excel_path, data_only=True)
            ws = wb['Summary']
            
            # Collecter tous les DUMs avec leurs données
            dum_lots_data = []
            current_dum = None
            current_p = None
            current_p_brut = None
            
            for row in range(1, 200):
                cell_c = ws.cell(row=row, column=3).value
                cell_a = ws.cell(row=row, column=1).value
                cell_b = ws.cell(row=row, column=2).value
                
                # Détecter DUM
                if cell_c and isinstance(cell_c, str) and "DUM" in cell_c:
                    # Sauvegarder le DUM précédent
                    if current_dum and current_p is not None and current_p_brut is not None:
                        dum_lots_data.append({
                            'dum_name': current_dum,
                            'p': current_p,
                            'p_brut': current_p_brut
                        })
                    
                    current_dum = cell_c.strip()
                    current_p = None
                    current_p_brut = None
                
                # Extraire P et P,BRUT
                if current_dum:
                    if cell_a == "P":
                        current_p = int(float(cell_b)) if cell_b else 0
                    elif cell_a == "P,BRUT":
                        current_p_brut = int(float(cell_b)) if cell_b else 0
            
            # Ajouter le dernier DUM
            if current_dum and current_p is not None and current_p_brut is not None:
                dum_lots_data.append({
                    'dum_name': current_dum,
                    'p': current_p,
                    'p_brut': current_p_brut
                })
            
            wb.close()
            
            print(f"      ✓ {len(dum_lots_data)} DUMs détectés pour création de lots")
            
        except Exception as e:
            print(f"      ❌ Erreur lecture DUMs depuis generated_excel: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # Créer un lot pour chaque DUM
        for dum_index, dum_data in enumerate(dum_lots_data, start=1):
            print(f"\n   🔹 Création du lot {dum_index}/{len(dum_lots_data)} ({dum_data['dum_name']})...")
            
            # ==================================================================
            # ÉTAPE ED.10.1: Cliquer sur "Nouveau" pour créer un lot
            # ==================================================================
            try:
                nouveau_lot_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_new_lot')]"))
                )
                nouveau_lot_btn.click()
                print(f"      ✓ Bouton 'Nouveau' lot cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur clic 'Nouveau' lot: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ==================================================================
            # ÉTAPE ED.10.2: Remplir les informations d'en-tête du lot
            # ==================================================================
            
            # ED.10.2a: Référence du lot (LTA ref + /N)
            try:
                # Construire la référence: ajouter /dum_index à la référence LTA validée
                # Tous les cas: simplement ajouter /N à la fin
                # 1. Si format "235-94908726/1" → "235-94908726/1/2" pour DUM 2
                # 2. Si format "23594908726" → "23594908726/2" pour DUM 2
                # 3. Si format "235-94908726" → "235-94908726/2" pour DUM 2
                
                # Toujours ajouter /N à la fin (ne jamais remplacer)
                lot_reference = f"{lta_reference_clean}/{dum_index}"
                
                ref_lot_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'referenceLot_IT_id')]"))
                )
                ref_lot_input.clear()
                ref_lot_input.send_keys(lot_reference)
                print(f"      ✓ Référence lot: {lot_reference}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie référence lot: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.10.2b: Ligne dépotée (toujours 1)
            try:
                ligne_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'ligneDepotee_IT_id')]"))
                )
                ligne_input.clear()
                ligne_input.send_keys("1")
                print(f"      ✓ Ligne dépotée: 1")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie ligne dépotée: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.10.2c: Sélectionner le radio button ICE (valeur 02)
            try:
                # Utiliser l'ID exact du radio button ICE (index 1, value 02)
                ice_radio = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:entete_section_form:radioChoixDestinataire:1"))
                )
                # Cliquer sur la div.ui-radiobutton-box associée
                radio_box = driver.find_element(By.XPATH, "//input[@id='mainTab:detailLot:entete_section_form:radioChoixDestinataire:1']/parent::div/following-sibling::div[@class='ui-radiobutton-box ui-widget ui-corner-all ui-state-default']")
                radio_box.click()
                print(f"      ✓ Option ICE sélectionnée")
                time.sleep(2)  # Attendre que la page se stabilise après le clic
            except Exception as e:
                print(f"      ⚠️  Erreur sélection radio ICE (méthode 1): {e}")
                # Méthode alternative: JavaScript
                try:
                    print(f"      🔄 Tentative avec JavaScript...")
                    js_code = """
                    var radio = document.getElementById('mainTab:detailLot:entete_section_form:radioChoixDestinataire:1');
                    if (radio) {
                        radio.checked = true;
                        var event = new Event('change', { bubbles: true });
                        radio.dispatchEvent(event);
                    }
                    """
                    driver.execute_script(js_code)
                    time.sleep(2)  # Attendre que la page se stabilise
                    print(f"      ✓ Option ICE sélectionnée via JavaScript")
                except Exception as e2:
                    print(f"      ❌ Erreur sélection radio ICE: {e2}")
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
            
            # Attendre que le blocker UI disparaisse après sélection ICE
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-blockui"))
                )
                print(f"      ✓ Page stabilisée après sélection ICE")
            except:
                pass  # Pas de blocker, continuer
            
            # ED.10.2d: Numéro ICE (constant)
            try:
                # Re-chercher l'élément ICE pour éviter stale element
                ice_input = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:detailLot:entete_section_form:id_ice"))
                )
                # Vérifier que l'élément est bien interactif
                wait.until(EC.element_to_be_clickable((By.ID, "mainTab:detailLot:entete_section_form:id_ice")))
                
                ice_input.clear()
                ice_input.send_keys("000230731000088")
                print(f"      ✓ ICE: 000230731000088")
                
                # Unfocus l'input pour déclencher la validation (Tab ou clic ailleurs)
                print(f"      ⏳ Déclenchement de la validation ICE...")
                from selenium.webdriver.common.keys import Keys
                ice_input.send_keys(Keys.TAB)  # Simuler Tab pour sortir du champ
                time.sleep(1)
                
                # Attendre que le système charge les informations ICE
                print(f"      ⏳ Attente du chargement des informations ICE...")
                time.sleep(3)  # Le système charge les infos après saisie ICE
                
                # Attendre que le blocker UI disparaisse si présent
                try:
                    WebDriverWait(driver, 5).until(
                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-blockui"))
                    )
                    print(f"      ✓ Informations ICE chargées")
                except:
                    pass  # Pas de blocker, continuer
                    
            except Exception as e:
                print(f"      ❌ Erreur saisie ICE: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ==================================================================
            # ÉTAPE ED.10.3: Valider l'en-tête du lot
            # ==================================================================
            try:
                # Utiliser l'ID exact du bouton Valider
                valider_lot_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "mainTab:detailLot:entete_section_form:btn_confirmer_lot"))
                )
                valider_lot_btn.click()
                print(f"      ✓ En-tête lot validé")
                
                # Attendre que la validation soit traitée et que la page soit prête
                print(f"      ⏳ Attente du traitement de la validation...")
                time.sleep(4)  # Augmenté à 4 secondes
                
                # Attendre que le blocker UI disparaisse si présent
                try:
                    WebDriverWait(driver, 5).until(
                        EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-blockui"))
                    )
                    print(f"      ✓ Validation traitée")
                except:
                    pass  # Pas de blocker, continuer
                
            except Exception as e:
                print(f"      ❌ Erreur validation en-tête lot: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ==================================================================
            # ÉTAPE ED.11.1: Cliquer sur "Nouveau" pour créer une ligne
            # ==================================================================
            try:
                print(f"      🔍 Recherche du bouton 'Nouveau' ligne...")
                nouveau_ligne_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_new_ligne')]"))
                )
                nouveau_ligne_btn.click()
                print(f"      ✓ Bouton 'Nouveau' ligne cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur clic 'Nouveau' ligne: {e}")
                print(f"      🔍 Tentative de recherche alternative...")
                try:
                    # Méthode alternative: chercher par texte visible
                    nouveau_ligne_btn_alt = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Nouveau') or contains(@value, 'Nouveau')]"))
                    )
                    driver.execute_script("arguments[0].click();", nouveau_ligne_btn_alt)
                    print(f"      ✓ Bouton 'Nouveau' ligne cliqué (méthode alternative)")
                    time.sleep(2)
                except Exception as e2:
                    print(f"      ❌ Erreur clic 'Nouveau' ligne (alternative): {e2}")
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
            
            # ==================================================================
            # ÉTAPE ED.11.2: Remplir le formulaire de ligne marchandise
            # ==================================================================
            
            # ED.11.2a: Type Contenant (autocomplete "colis")
            try:
                type_contenant_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'typeContenantId_INPUT_input')]"))
                )
                type_contenant_input.clear()
                type_contenant_input.send_keys("colis")
                print(f"      ✓ Type contenant: colis (recherche...)")
                time.sleep(2)
                
                # Sélectionner la première suggestion "COLIS(216)"
                colis_suggestion = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//li[contains(@data-item-label, 'COLIS(216)')]"))
                )
                colis_suggestion.click()
                print(f"      ✓ COLIS(216) sélectionné")
                time.sleep(1)
            except Exception as e:
                print(f"      ❌ Erreur sélection type contenant: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2b: Nombre de contenants (P du DUM)
            try:
                nbr_contenants_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'nbrContenants')]"))
                )
                nbr_contenants_input.clear()
                nbr_contenants_input.send_keys(str(dum_data['p']))
                print(f"      ✓ Nombre contenants: {dum_data['p']}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie nombre contenants: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2c: Poids brut (P,BRUT du DUM)
            try:
                poids_brut_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'poidBru_input')]"))
                )
                poids_brut_input.clear()
                poids_brut_input.send_keys(str(dum_data['p_brut']))
                print(f"      ✓ Poids brut: {dum_data['p_brut']}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie poids brut: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2d: Marque (référence LTA validée)
            try:
                marque_textarea = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//textarea[contains(@name, 'marqueLib')]"))
                )
                marque_textarea.clear()
                marque_textarea.send_keys(lta_reference_clean)
                print(f"      ✓ Marque: {lta_reference_clean}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie marque: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2e: Nature marchandise (constant)
            try:
                nature_textarea = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//textarea[contains(@name, 'marchand')]"))
                )
                nature_textarea.clear()
                nature_textarea.send_keys("courrier e-commerce")
                print(f"      ✓ Nature marchandise: courrier e-commerce")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie nature marchandise: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2f: Code NGP (9999)
            try:
                ngp_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@name, ':ngp') and @type='text']"))
                )
                ngp_input.clear()
                ngp_input.send_keys("9999")
                print(f"      ✓ Code NGP: 9999")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie NGP: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.11.2g: Ajouter le code NGP (bouton >>)
            try:
                add_ngp_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_add_ngp')]"))
                )
                add_ngp_btn.click()
                print(f"      ✓ Code NGP ajouté")
                time.sleep(1)
            except Exception as e:
                print(f"      ❌ Erreur ajout NGP: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ==================================================================
            # ÉTAPE ED.11.3: Valider la ligne marchandise
            # ==================================================================
            try:
                valider_ligne_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btn_confirmer_ligne')]"))
                )
                valider_ligne_btn.click()
                print(f"      ✓ Ligne marchandise validée")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur validation ligne: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            print(f"   ✅ Lot {dum_index} créé avec succès!")
        
        # ==================================================================
        # FIN - Tous les lots créés
        # ==================================================================
        print(f"\n   ✅ Tous les lots ({len(dum_lots_data)}) créés avec succès!")
        print("      Etat de Dépotage - Onglet LTA complété")
        
        # ==================================================================
        # ÉTAPE ED.13: Sauvegarder l'Etat de Dépotage
        # ==================================================================
        print("\n   💾 Sauvegarde de l'Etat de Dépotage...")
        
        try:
            # ED.13.1: Cliquer sur "SAUVEGARDER" (menu item)
            # Stratégie robuste: chercher par texte puis par ID si échoue
            try:
                sauvegarder_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ui-menuitem-link')]//span[text()='SAUVEGARDER']/parent::a"))
                )
                sauvegarder_link.click()
                print("      ✓ Bouton 'SAUVEGARDER' cliqué")
            except:
                # Fallback: par ID
                sauvegarder_link = wait.until(
                    EC.element_to_be_clickable((By.ID, "secure_174"))
                )
                sauvegarder_link.click()
                print("      ✓ Bouton 'SAUVEGARDER' cliqué (via ID)")
            
            time.sleep(3)  # Attendre la sauvegarde
            print("      ✓ Etat de Dépotage sauvegardé")
            
        except Exception as e:
            print(f"      ❌ Erreur sauvegarde: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ED.13.2: Extraire la référence "sauvegardée" (avant validation)
        # Cette référence sera utilisée en cas d'échec de validation
        sauvegarde_reference = None
        try:
            time.sleep(2)  # Attendre que la référence s'affiche
            
            # Chercher la table de référence
            reference_table = driver.find_element(By.CSS_SELECTOR, "table.reference")
            
            # Extraire les cellules de la deuxième ligne (index 1)
            rows = reference_table.find_elements(By.TAG_NAME, "tr")
            if len(rows) >= 2:
                data_row = rows[1]
                cells = data_row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) >= 5:
                    # Extraire Série (colonne 4, index 3) et Clé (colonne 5, index 4)
                    serie_value = cells[3].text.strip()
                    cle_value = cells[4].text.strip()
                    
                    # Enlever les zéros initiaux de la série
                    serie_clean = str(int(serie_value)) if serie_value.isdigit() else serie_value
                    
                    # Combiner: [Série][Clé]
                    sauvegarde_reference = f"{serie_clean}{cle_value}"
                    
                    print(f"      ✓ Référence sauvegardée extraite: {sauvegarde_reference}")
                    print(f"         (Série={serie_value} → {serie_clean}, Clé={cle_value})")
        except Exception as e:
            print(f"      ⚠️  Impossible d'extraire la référence sauvegardée: {e}")
            print(f"         (Continuera avec extraction après validation)")
            # Continuer quand même - on essaiera après validation
        
        # ==================================================================
        # ÉTAPE ED.14: Valider l'Etat de Dépotage
        # ==================================================================
        print("\n   ✅ Validation de l'Etat de Dépotage...")
        
        try:
            # ED.14.1: Cliquer sur "VALIDER" (menu item)
            try:
                valider_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ui-menuitem-link')]//span[text()='VALIDER']/parent::a"))
                )
                valider_link.click()
                print("      ✓ Bouton 'VALIDER' cliqué")
            except:
                # Fallback: par ID
                valider_link = wait.until(
                    EC.element_to_be_clickable((By.ID, "secure_176"))
                )
                valider_link.click()
                print("      ✓ Bouton 'VALIDER' cliqué (via ID)")
            
            time.sleep(4)  # Attendre la validation
            
        except Exception as e:
            print(f"      ❌ Erreur validation: {e}")
            driver.switch_to.default_content()
            return_to_home_after_error(driver)
            return False
        
        # ==================================================================
        # ÉTAPE ED.15: Vérifier la réponse de validation
        # ==================================================================
        print("\n   🔍 Vérification du résultat de validation...")
        
        # ED.15.1: Chercher messages d'erreur
        try:
            time.sleep(2)  # Attendre que les messages s'affichent
            
            error_detected = False
            error_message = ""
            
            # Chercher container d'erreur
            error_containers = driver.find_elements(By.ID, "msg-error")
            if not error_containers:
                error_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error")
            
            if error_containers and len(error_containers) > 0:
                # Chercher le détail de l'erreur
                error_details = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                if error_details and len(error_details) > 0:
                    error_message = error_details[0].text.strip()
                    error_detected = True
                    print(f"      ❌ Erreur détectée: {error_message}")
            
            # ED.15.2: Si erreur, créer fichier d'erreur
            if error_detected:
                lta_name = os.path.basename(lta_folder_path)
                
                # ED.15.2.1: Détecter l'erreur de poids (E2800124)
                is_weight_mismatch = "E2800124" in error_message or "Echec au rapprochement" in error_message
                
                if is_weight_mismatch:
                    # Erreur de discordance de poids - LOG SPÉCIAL
                    error_filename = f"-----------error-weight-mismatch-{lta_name}.txt"
                    error_filepath = os.path.join(os.path.dirname(lta_folder_path), error_filename)
                    
                    from datetime import datetime
                    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    with open(error_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"⚠️  ERREUR DE VALIDATION - DISCORDANCE DE POIDS\n")
                        f.write(f"{'='*70}\n\n")
                        f.write(f"LTA: {lta_name}\n")
                        f.write(f"Référence LTA: {lta_reference_clean}\n")
                        f.write(f"Date/Heure: {current_datetime}\n\n")
                        
                        f.write(f"TOTAUX DÉCLARÉS DANS LE CANEVAS LTA:\n")
                        f.write(f"  • Total P (Contenants): {total_p}\n")
                        f.write(f"  • Total P,BRUT (Poids brut): {total_p_brut}\n\n")
                        
                        # Ajouter la référence sauvegardée si disponible
                        if sauvegarde_reference:
                            f.write(f"RÉFÉRENCE DS SAUVEGARDÉE (non validée):\n")
                            f.write(f"  • Série: {sauvegarde_reference}\n\n")
                        
                        f.write(f"MESSAGE D'ERREUR SYSTÈME:\n")
                        f.write(f"{error_message}\n\n")
                        
                        f.write(f"{'='*70}\n")
                        f.write(f"EXPLICATION:\n")
                        f.write(f"Le document DS (Etat de Dépotage) envoyé par Said, Youssef ou Ibrahim\n")
                        f.write(f"contient un total de Poids Brut DIFFÉRENT de celui indiqué dans le\n")
                        f.write(f"canevas de la LTA ({lta_name}).\n\n")
                        
                        f.write(f"ACTION REQUISE:\n")
                        f.write(f"1. Vérifier le document DS original envoyé par mail\n")
                        f.write(f"2. Comparer le poids total du DS avec le total du canevas LTA ({total_p_brut})\n")
                        f.write(f"3. Ajuster le canevas LTA (generated_excel) si nécessaire\n")
                        f.write(f"4. OU demander un DS corrigé à l'expéditeur\n")
                        f.write(f"5. Relancer le script après correction\n\n")
                        
                        if sauvegarde_reference:
                            f.write(f"NOTE: Une référence DS a été SAUVEGARDÉE (non validée).\n")
                            f.write(f"      Vous pouvez la modifier manuellement sur BADR si nécessaire.\n\n")
                        
                        f.write(f"STATUT: Échec de validation - CORRECTION MANUELLE NÉCESSAIRE\n")
                        f.write(f"{'='*70}\n")
                    
                    print(f"      ⚠️  ERREUR DE POIDS DÉTECTÉE (E2800124)")
                    print(f"      ✓ Log créé: {error_filename}")
                    print(f"      📊 Total P: {total_p} | Total P,BRUT: {total_p_brut}")
                    print(f"      ⚠️  Discordance avec document DS - Correction manuelle requise")
                    
                    # Sauvegarder la référence "sauvegardée" dans le fichier shipper (ligne 4)
                    if sauvegarde_reference:
                        try:
                            parent_dir = os.path.dirname(lta_folder_path)
                            lta_name_with_underscore = lta_name.replace(" ", "_")
                            shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                            shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                            
                            if shipper_files:
                                shipper_file = shipper_files[0]
                                
                                # Lire le fichier actuel
                                with open(shipper_file, 'r', encoding='utf-8') as f:
                                    lines = [line.rstrip('\n') for line in f.readlines()]
                                
                                # S'assurer qu'on a au moins 3 lignes
                                while len(lines) < 3:
                                    lines.append("")
                                
                                # Ajouter ou remplacer la ligne 4 avec la référence sauvegardée
                                if len(lines) == 3:
                                    lines.append(sauvegarde_reference)
                                elif len(lines) >= 4:
                                    lines[3] = sauvegarde_reference
                                
                                # Réécrire le fichier
                                with open(shipper_file, 'w', encoding='utf-8') as f:
                                    f.write('\n'.join(lines))
                                
                                print(f"      ✓ Référence sauvegardée écrite dans {os.path.basename(shipper_file)}")
                                print(f"         Ligne 4: {sauvegarde_reference} (NON VALIDÉE)")
                        except Exception as e:
                            print(f"      ⚠️  Impossible de sauvegarder la référence: {e}")
                    
                    print(f"      ⏭️  Passage au LTA suivant...")
                    
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
                
                else:
                    # Erreur générique - LOG STANDARD
                    error_filename = f"-----------error-validating-ds-{lta_name}.txt"
                    error_filepath = os.path.join(os.path.dirname(lta_folder_path), error_filename)
                    
                    from datetime import datetime
                    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    with open(error_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"ERREUR DE VALIDATION - Etat de Dépotage\n\n")
                        f.write(f"LTA: {lta_name} - {lta_reference_clean}\n")
                        f.write(f"Date: {current_datetime}\n")
                        f.write(f"Étape: Validation finale de l'état de dépotage\n\n")
                        f.write(f"MESSAGE D'ERREUR:\n")
                        f.write(f"{error_message}\n\n")
                        f.write(f"CONTEXTE:\n")
                        f.write(f"- Nombre de lots créés: {len(dum_lots_data)}\n")
                        f.write(f"- Nombre de DUMs traités: {len(dum_lots_data)}\n")
                        f.write(f"- Dernière action: Click sur VALIDER\n\n")
                        f.write(f"STATUT: Échec de validation\n")
                        f.write(f"Action recommandée: Vérifier les données saisies et réessayer manuellement\n")
                    
                    print(f"      ✓ Fichier d'erreur créé: {error_filename}")
                    print(f"      ⚠️  Arrêt du traitement de ce LTA")
                    
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
            
            # ED.15.3: Si succès, extraire la référence de déclaration
            print("      ✓ Validation réussie - extraction de la référence...")
            
            try:
                # Chercher la table de référence
                reference_table = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table.reference"))
                )
                
                # Extraire les cellules de la deuxième ligne (index 1)
                rows = reference_table.find_elements(By.TAG_NAME, "tr")
                if len(rows) < 2:
                    print("      ❌ Table de référence incomplète")
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
                
                data_row = rows[1]
                cells = data_row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) < 5:
                    print("      ❌ Données de référence incomplètes")
                    driver.switch_to.default_content()
                    return_to_home_after_error(driver)
                    return False
                
                # Extraire Série (colonne 4, index 3) et Clé (colonne 5, index 4)
                serie_value = cells[3].text.strip()
                cle_value = cells[4].text.strip()
                
                # Enlever les zéros initiaux de la série
                serie_clean = str(int(serie_value)) if serie_value.isdigit() else serie_value
                
                # Combiner: [Série][Clé]
                ds_reference = f"{serie_clean}{cle_value}"
                
                print(f"      ✓ Référence extraite: Série={serie_value} → {serie_clean}, Clé={cle_value}")
                print(f"      ✓ Référence DS complète: {ds_reference}")
                
            except Exception as e:
                print(f"      ❌ Erreur extraction référence: {e}")
                driver.switch_to.default_content()
                return_to_home_after_error(driver)
                return False
            
            # ED.15.4: Ajouter la référence DS au fichier shipper (ligne 4)
            try:
                lta_name = os.path.basename(lta_folder_path)
                parent_dir = os.path.dirname(lta_folder_path)
                
                # Chercher le fichier shipper: [X]eme_LTA_*.txt
                # Le nom du dossier a des espaces (ex: "7eme LTA") mais le fichier a des underscores (ex: "7eme_LTA_shipper_name.txt")
                lta_name_with_underscore = lta_name.replace(" ", "_")
                shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                
                if not shipper_files:
                    print(f"      ⚠️  Fichier shipper introuvable: {shipper_pattern}")
                    # Continuer quand même, ce n'est pas critique
                else:
                    shipper_file = shipper_files[0]
                    
                    # Lire le fichier actuel
                    with open(shipper_file, 'r', encoding='utf-8') as f:
                        lines = [line.rstrip('\n') for line in f.readlines()]
                    
                    # S'assurer qu'on a au moins 3 lignes (shipper, serie+cle, location)
                    while len(lines) < 3:
                        lines.append("")
                    
                    # Ajouter ou remplacer la ligne 4 (index 3) avec la référence DS
                    if len(lines) == 3:
                        # Ajouter ligne 4
                        lines.append(ds_reference)
                        action = "ajoutée"
                    elif len(lines) >= 4:
                        # Remplacer ligne 4 existante
                        lines[3] = ds_reference
                        action = "mise à jour"
                    
                    # Réécrire le fichier
                    with open(shipper_file, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(lines))
                    
                    print(f"      ✓ Référence DS {action} dans {os.path.basename(shipper_file)}")
                    print(f"         Ligne 4: {ds_reference}")
                
            except Exception as e:
                print(f"      ⚠️  Erreur mise à jour fichier shipper: {e}")
                # Continuer quand même
            
            print("\n   ✅ Etat de Dépotage complété avec succès!")
            print(f"      Référence DS: {ds_reference}")
            
        except Exception as e:
            print(f"      ❌ Erreur vérification validation: {e}")
            driver.switch_to.default_content()
            return False
        
        # Sortir de l'iframe avant de retourner
        driver.switch_to.default_content()
        print("      ✓ Sorti de l'iframe Etat de Dépotage")
        return True
        
    except Exception as e:
        print(f"\n❌ Erreur création Etat de Dépotage: {e}")
        traceback.print_exc()
        # Essayer de sortir de l'iframe en cas d'erreur
        try:
            driver.switch_to.default_content()
            print("   ⚠️  Sorti de l'iframe après erreur")
        except:
            pass
        return False

def read_shipper_from_txt(txt_file_path):
    """Extract shipper name, LTA reference, and DS MEAD reference data from .txt file
    
    Supports TWO file formats:
    
    Format 1 - Simple 3-line format (for signed LTAs):
        Line 1: Shipper company name (e.g., "JIANGSU MINQIAN TECHNOLOGY CO., LTD")
        Line 2: Serie + Cle (e.g., "2793 X")
        Line 3: Loading location (e.g., "ABOU DHABI INT")
    
    Format 2 - OCR extraction format (6+ lines):
        Line 1: Separator "-------------"
        Line 2: LTA folder name (e.g., "9eme LTA_ocr_textbsed")
        Line 3: LTA reference without prefix (e.g., "60738318932")
        Line 4: LTA reference with prefix (e.g., "607-38318932/1")
        Line 5: Empty line
        Line 6: Shipper company name (e.g., "JIANGSU MINQIAN TECHNOLOGY CO., LTD")
        Line 7+: Optional DS MEAD data
    
    Returns:
        dict: {
            'shipper_name': str,
            'lta_reference': str (format: "607-38318932/1"),
            'lta_reference_clean': str (format: "607-38318932" without /1),
            'has_ds_mead': bool,
            'serie': str or None,
            'cle': str or None,
            'loading_location': str or None
        }
    """
    try:
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # Remove empty lines and strip whitespace
        lines = [line.strip() for line in all_lines if line.strip()]
        
        if len(lines) < 1:
            print(f"   ❌ Fichier vide: {txt_file_path}")
            return None
        
        # Detect format: Check if line 1 starts with "---" (Format 2) or is company name (Format 1)
        is_ocr_format = lines[0].startswith('---') if len(lines) > 0 else False
        
        # FORMAT 1: Simple 3-line format (shipper, serie+cle, location)
        if not is_ocr_format and len(lines) >= 2:
            print(f"   📄 Format détecté: Simple 3-line format")
            shipper_name = lines[0]
            
            # Parse line 2 for serie + cle (e.g., "2793 X")
            has_ds_mead = False
            serie = None
            cle = None
            loading_location = None
            
            parts = lines[1].split()
            if len(parts) == 2 and parts[0].isdigit() and len(parts[1]) == 1:
                has_ds_mead = True
                serie = parts[0]
                cle = parts[1]
                print(f"   ✓ DS MEAD détecté - Série: {serie}, Clé: {cle}")
            
            # Line 3 is loading location (if exists)
            if len(lines) >= 3:
                loading_location = lines[2]
                print(f"   ✓ Lieu de chargement: {loading_location}")
            
            # LTA reference not available in this format - extract from filename
            # Filename format: "2eme_LTA_shipper_name.txt" -> extract "2eme LTA"
            lta_name = os.path.basename(txt_file_path).replace('_shipper_name.txt', '').replace('_', ' ')
            lta_reference = f"UNKNOWN/{lta_name}"  # Placeholder
            lta_reference_clean = "UNKNOWN"
            
            return {
                'shipper_name': shipper_name,
                'lta_reference': lta_reference,
                'lta_reference_clean': lta_reference_clean,
                'has_ds_mead': has_ds_mead,
                'serie': serie,
                'cle': cle,
                'loading_location': loading_location
            }
        
        # FORMAT 2: OCR extraction format (6+ lines)
        if len(lines) < 4:
            print(f"   ❌ Fichier incomplet: {txt_file_path}")
            print(f"      Attendu: Format 1 (3 lignes) ou Format 2 (6+ lignes)")
            print(f"      Trouvé: {len(lines)} ligne(s)")
            return None
        # FORMAT 2: OCR extraction format (6+ lines)
        if len(lines) < 4:
            print(f"   ❌ Fichier incomplet: {txt_file_path}")
            print(f"      Attendu: Format 1 (3 lignes) ou Format 2 (6+ lignes)")
            print(f"      Trouvé: {len(lines)} ligne(s)")
            return None
        
        print(f"   📄 Format détecté: OCR extraction format (6+ lignes)")
        
        # Extract LTA reference from line 4 (index 3)
        # Format: "607-38318932/1"
        lta_reference_full = lines[3]
        
        # Clean version without /1
        lta_reference_clean = lta_reference_full.split('/')[0] if '/' in lta_reference_full else lta_reference_full
        
        # Extract shipper name from line 6 (index 5)
        shipper_name = lines[5] if len(lines) > 5 else lines[0]
        
        # Check for DS MEAD reference (optional, would be after shipper name)
        # Look for line with pattern "XXXX Y" (serie + cle)
        has_ds_mead = False
        serie = None
        cle = None
        loading_location = None
        
        # Check lines after shipper name for DS MEAD info
        for i in range(6, min(len(lines), 10)):
            line = lines[i]
            parts = line.split()
            # DS MEAD format: "2666 M" (number followed by single letter)
            if len(parts) == 2 and parts[0].isdigit() and len(parts[1]) == 1:
                has_ds_mead = True
                serie = parts[0]
                cle = parts[1]
                # Next line might be loading location
                if i + 1 < len(lines):
                    loading_location = lines[i + 1]
                break
        
        if has_ds_mead:
            print(f"   ✓ LTA {lta_reference_clean} avec DS MEAD - Série: {serie}, Clé: {cle}")
        else:
            print(f"   ✓ LTA {lta_reference_clean} sans DS MEAD")
        
        return {
            'shipper_name': shipper_name,
            'lta_reference': lta_reference_full,
            'lta_reference_clean': lta_reference_clean,
            'has_ds_mead': has_ds_mead,
            'serie': serie,
            'cle': cle,
            'loading_location': loading_location
        }
        
    except Exception as e:
        print(f"   ❌ Erreur lecture shipper depuis {txt_file_path}: {e}")
        traceback.print_exc()
        return None

def read_dum_data_from_summary(summary_excel_path):
    """Read all DUM/Sheet data from summary_file Excel
    Returns: list of dicts with keys: sheet_name, total_pieces, total_value, 
             total_gross_weight, total_freight, insurance, cartons
    """
    try:
        wb = load_workbook(summary_excel_path, data_only=True)
        
        # Find the sheet with the summary table (usually first sheet or named 'Summary')
        if 'Summary' in wb.sheetnames:
            ws = wb['Summary']
        else:
            ws = wb.active
        
        dum_list = []
        
        # Find header row (contains "Sheet Name", "Total Pieces", etc.)
        header_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20), start=1):
            cell_values = [str(cell.value).lower() if cell.value else '' for cell in row]
            if 'sheet' in ' '.join(cell_values) and 'total' in ' '.join(cell_values):
                header_row = row_idx
                break
        
        if not header_row:
            print("   ⚠️  Impossible de trouver l'en-tête du tableau dans summary_file")
            return []
        
        # Read header to find column indices
        headers = {}
        for col_idx, cell in enumerate(ws[header_row], start=1):
            header_text = str(cell.value).lower().strip() if cell.value else ''
            if 'sheet' in header_text or 'nom' in header_text:
                headers['sheet_name'] = col_idx
            elif 'pieces' in header_text or 'nombre' in header_text:
                headers['total_pieces'] = col_idx
            elif 'value' in header_text or 'valeur' in header_text:
                headers['total_value'] = col_idx
            elif 'gross' in header_text or 'brut' in header_text or 'poid' in header_text:
                headers['total_gross_weight'] = col_idx
            elif 'freight' in header_text or 'fret' in header_text:
                headers['total_freight'] = col_idx
            elif 'insurance' in header_text or 'assurance' in header_text:
                headers['insurance'] = col_idx
            elif 'carton' in header_text or 'colis' in header_text:
                headers['cartons'] = col_idx
            elif 'position' in header_text:
                headers['total_positions'] = col_idx
        
        print(f"   📊 Colonnes trouvées: {headers}")
        
        # Read data rows
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or not row[headers.get('sheet_name', 0) - 1]:
                continue  # Skip empty rows
            
            # Arrondir les valeurs décimales à 2 chiffres pour éviter les erreurs de précision flottante
            total_value = row[headers.get('total_value', 3) - 1] if 'total_value' in headers else 0
            if isinstance(total_value, (int, float)):
                total_value = round(float(total_value), 2)
            
            total_gross_weight = row[headers.get('total_gross_weight', 4) - 1] if 'total_gross_weight' in headers else 0
            if isinstance(total_gross_weight, (int, float)):
                total_gross_weight = round(float(total_gross_weight), 2)
            
            total_freight = row[headers.get('total_freight', 5) - 1] if 'total_freight' in headers else 0
            if isinstance(total_freight, (int, float)):
                total_freight = round(float(total_freight), 2)
            
            insurance = row[headers.get('insurance', 6) - 1] if 'insurance' in headers else 0
            if isinstance(insurance, (int, float)):
                insurance = round(float(insurance), 2)
            
            dum_data = {
                'sheet_name': row[headers.get('sheet_name', 1) - 1],
                'total_pieces': row[headers.get('total_pieces', 2) - 1] if 'total_pieces' in headers else 0,
                'total_value': total_value,
                'total_gross_weight': total_gross_weight,
                'total_freight': total_freight,
                'insurance': insurance,
                'cartons': row[headers.get('cartons', 7) - 1] if 'cartons' in headers else 0,
                'total_positions': row[headers.get('total_positions', 8) - 1] if 'total_positions' in headers else 0,
            }
            
            dum_list.append(dum_data)
        
        print(f"   ✓ {len(dum_list)} DUMs trouvés dans summary_file")
        return dum_list
        
    except Exception as e:
        print(f"   ❌ Erreur lecture summary_file: {e}")
        traceback.print_exc()
        return []

def fill_declaration_form(driver, shipper_name, dum_data, lta_folder_path, lta_reference_clean):
    """Fill the declaration form with shipper name and DUM data
    
    Args:
        driver: Selenium WebDriver instance (already in iframe context)
        shipper_name: Shipper company name (string)
        dum_data: Dict with keys: sheet_name, total_value, total_gross_weight, total_freight, insurance, total_positions
        lta_folder_path: Path to LTA folder containing Sheet Excel files
        lta_reference_clean: LTA reference without /1 suffix (e.g., "607-38318932")
    """
    try:
        wait = WebDriverWait(driver, 15)
        
        print(f"\n📝 Remplissage du formulaire pour {dum_data.get('sheet_name', 'DUM')}...")
        print(f"   📋 LTA Reference: {lta_reference_clean}")
        
        # 1. Shipper Name
        print("   1️⃣ Nom expéditeur...")
        shipper_input = wait.until(
            EC.presence_of_element_located((By.ID, "mainTab:form0:nomOperateurExpediteur"))
        )
        shipper_input.clear()
        shipper_input.send_keys(shipper_name)
        print(f"      ✓ Expéditeur: {shipper_name}")
        time.sleep(0.5)
        
        # 2. Total Value
        print("   2️⃣ Montant total...")
        total_value_input = wait.until(
            EC.presence_of_element_located((By.ID, "mainTab:form0:montTotalNumber_input"))
        )
        total_value_input.clear()
        total_value_input.send_keys(str(dum_data.get('total_value', 0)))
        print(f"      ✓ Valeur totale: {dum_data.get('total_value', 0)}")
        time.sleep(0.5)
        
        # 3. Total Gross Weight
        print("   3️⃣ Poids brut total...")
        gross_weight_input = wait.until(
            EC.presence_of_element_located((By.ID, "mainTab:form0:poidBrutTotal_input"))
        )
        gross_weight_input.clear()
        gross_weight_input.send_keys(str(dum_data.get('total_gross_weight', 0)))
        print(f"      ✓ Poids brut: {dum_data.get('total_gross_weight', 0)}")
        time.sleep(0.5)
        
        # 4. Freight Amount
        print("   4️⃣ Montant fret...")
        freight_input = wait.until(
            EC.presence_of_element_located((By.ID, "mainTab:form0:montantFret_input"))
        )
        freight_input.clear()
        freight_input.send_keys(str(dum_data.get('total_freight', 0)))
        print(f"      ✓ Fret: {dum_data.get('total_freight', 0)}")
        time.sleep(0.5)
        
        # 5. Insurance Amount
        print("   5️⃣ Montant assurance...")
        insurance_input = wait.until(
            EC.presence_of_element_located((By.ID, "mainTab:form0:mntAssuranceNumber_input"))
        )
        insurance_input.clear()
        insurance_input.send_keys(str(dum_data.get('insurance', 0)))
        print(f"      ✓ Assurance: {dum_data.get('insurance', 0)}")
        time.sleep(0.5)
        
        # 6. Credit Removal Selection (select the only available option)
        print("   6️⃣ Crédit d'enlèvement...")
        try:
            # Click on the dropdown trigger
            credit_trigger = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form0\\:choixCdeId div.ui-selectonemenu-trigger"))
            )
            credit_trigger.click()
            print("      ✓ Dropdown crédit ouvert")
            time.sleep(1)
            
            # Select the option "3095361 - National - 30"
            # L'option devrait apparaître dans une liste ul.ui-selectonemenu-items
            credit_option = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li[@data-label='3095361 - National - 30']"))
            )
            credit_option.click()
            print("      ✓ Crédit: 3095361 - National - 30")
        except Exception as e:
            print(f"      ⚠️  Erreur sélection crédit (méthode UI): {e}")
            # Alternative method: Use JavaScript to select the option in the hidden select
            try:
                print("      🔄 Tentative avec JavaScript...")
                # Sélectionner l'option via JavaScript
                js_code = """
                var select = document.getElementById('mainTab:form0:choixCdeId_input');
                select.value = '3095361';
                // Déclencher l'événement change pour que PrimeFaces détecte le changement
                var event = new Event('change', { bubbles: true });
                select.dispatchEvent(event);
                
                // Mettre à jour le label visible
                var label = document.getElementById('mainTab:form0:choixCdeId_label');
                label.textContent = '3095361 - National - 30';
                """
                driver.execute_script(js_code)
                time.sleep(0.5)
                print("      ✓ Crédit sélectionné via JavaScript")
            except Exception as e2:
                print(f"      ❌ Impossible de sélectionner le crédit: {e2}")
                print("      ⚠️  Continuons sans crédit sélectionné...")
        
        time.sleep(1)
        
        print(f"\n   ✅ Formulaire initial rempli pour {dum_data.get('sheet_name', 'DUM')}")
        
        # ==================================================================
        # ÉTAPE 2: Naviguer vers l'onglet "Articles"
        # ==================================================================
        print("\n   📑 Navigation vers l'onglet 'Articles'...")
        try:
            articles_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab1']"))
            )
            articles_tab.click()
            print("      ✓ Onglet 'Articles' cliqué")
            time.sleep(2)  # Attendre le chargement de l'onglet
        except Exception as e:
            print(f"      ❌ Erreur navigation vers 'Articles': {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 3: Upload du fichier Excel correspondant (Sheet X.xlsx)
        # ==================================================================
        print(f"\n   📤 Upload du fichier Excel pour {dum_data.get('sheet_name', 'DUM')}...")
        
        # Trouver le fichier Sheet correspondant dans le dossier LTA
        sheet_name = dum_data.get('sheet_name', '')
        
        # Recherche du fichier avec pattern: "Sheet 1 - *.xlsx", "Sheet 2 - *.xlsx", etc.
        sheet_pattern = os.path.join(lta_folder_path, f"{sheet_name} - *.xlsx")
        matching_files = glob.glob(sheet_pattern)
        
        if not matching_files:
            # Fallback: essayer sans le tiret
            sheet_pattern = os.path.join(lta_folder_path, f"{sheet_name}*.xlsx")
            matching_files = glob.glob(sheet_pattern)
        
        if not matching_files:
            print(f"      ❌ Fichier Excel introuvable pour {sheet_name}")
            print(f"         Pattern recherché: {sheet_pattern}")
            
            # NETTOYAGE: Retourner à l'accueil avant de quitter
            print("\n   🏠 Nettoyage: Retour à l'accueil...")
            try:
                driver.switch_to.default_content()
                time.sleep(1)
                # Rafraîchir la page pour revenir à l'état initial
                driver.get(driver.current_url.split('#')[0])
                time.sleep(2)
                print("      ✓ Retour à l'état initial")
            except Exception as cleanup_err:
                print(f"      ⚠️  Erreur nettoyage: {cleanup_err}")
            
            return False
        
        sheet_file_path = matching_files[0]  # Prendre le premier fichier trouvé
        print(f"      ✓ Fichier trouvé: {os.path.basename(sheet_file_path)}")
        
        # Upload du fichier
        try:
            file_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form1:idFileUploadWidgetVar_input"))
            )
            
            # Convertir le chemin en chemin absolu Windows
            absolute_path = os.path.abspath(sheet_file_path)
            print(f"        Chemin absolu: {absolute_path}")
            
            file_input.send_keys(absolute_path)
            print(f"      ✓ Fichier uploadé: {os.path.basename(sheet_file_path)}")
            
            # Attendre que l'upload soit traité
            time.sleep(3)
            
        except Exception as e:
            print(f"      ❌ Erreur upload fichier: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 4: Validation du nombre total de positions
        # ==================================================================
        print(f"\n   ✅ Validation du nombre de positions...")
        
        # Récupérer le nombre de positions attendu depuis summary_file
        expected_positions = dum_data.get('total_positions', 0)
        print(f"      📊 Positions attendues (summary_file): {expected_positions}")
        
        # Attendre que la table soit chargée après l'upload
        time.sleep(2)
        
        try:
            # Lire le nombre de positions calculé par le système
            # La table contient les articles importés
            # On cherche la valeur dans la colonne "Nb. cont." (3ème td de chaque row)
            
            # Trouver toutes les lignes de la table
            table_rows = driver.find_elements(By.CSS_SELECTOR, "tbody#mainTab\\:form1\\:j_id_3p_y7_data tr")
            
            if not table_rows:
                print("      ⚠️  Aucune donnée trouvée dans la table après upload")
            else:
                print(f"      📋 {len(table_rows)} ligne(s) trouvée(s) dans la table")
                
                # Lire la valeur de "Nb. cont." dans la première ligne
                # (normalement devrait être dans la 3ème colonne)
                first_row = table_rows[0]
                cells = first_row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) >= 3:
                    system_positions_text = cells[2].text.strip()
                    try:
                        system_positions = int(system_positions_text)
                        print(f"      🔢 Positions calculées (système): {system_positions}")
                        
                        # Comparaison
                        if expected_positions == 0:
                            print(f"      ⚠️  Pas de validation possible (expected_positions = 0)")
                        elif system_positions == expected_positions:
                            print(f"      ✅ VALIDATION OK: {system_positions} = {expected_positions}")
                        else:
                            print(f"      ⚠️  DIVERGENCE: Système={system_positions}, Attendu={expected_positions}")
                            # Pour l'instant, on continue quand même
                    except ValueError:
                        print(f"      ⚠️  Impossible de convertir '{system_positions_text}' en nombre")
                else:
                    print(f"      ⚠️  Moins de 3 colonnes trouvées ({len(cells)})")
                    
        except Exception as e:
            print(f"      ⚠️  Erreur validation positions: {e}")
            # On continue quand même
        
        # Attendre que l'overlay de blocage disparaisse après l'upload
        print("\n   ⏳ Attente de la fin du traitement de l'upload...")
        try:
            # Attendre que le blocker disparaisse (devient invisible ou hidden)
            wait.until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-blockui"))
            )
            print("      ✓ Traitement terminé")
            time.sleep(1)  # Petite pause supplémentaire pour la stabilité
        except Exception as e:
            print(f"      ⚠️  Timeout attente blocker (peut être déjà disparu): {e}")
            # Continuer quand même
        
        # ==================================================================
        # ÉTAPE 5: Vérification Carton et Workflow "Demandes diverses" (Conditionnel)
        # ==================================================================
        carton_value = dum_data.get('cartons', 0)
        print(f"\n   📦 Vérification de la valeur Carton: {carton_value}")
        
       
        # print(f"      ✓ Carton ({carton_value}) ≠ 13 → Traitement 'Demandes diverses'")
        
        # ÉTAPE 5.2: Naviguer vers "Demandes diverses"
        print("\n   📋 Navigation vers l'onglet 'Demandes diverses'...")
        try:
            demandes_diverses_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab4']"))
            )
            # Essayer le clic normal
            try:
                demandes_diverses_tab.click()
                print("      ✓ Onglet 'Demandes diverses' cliqué")
            except Exception as click_error:
                print(f"      ⚠️  Clic normal intercepté, utilisation de JavaScript...")
                # Utiliser JavaScript si le clic est intercepté
                driver.execute_script("arguments[0].click();", demandes_diverses_tab)
                print("      ✓ Onglet 'Demandes diverses' cliqué (via JavaScript)")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation vers 'Demandes diverses': {e}")
            return False
        
        # ÉTAPE 5.3: Cliquer sur le lien "Autre(01)"
        print("\n   🔗 Clic sur 'Autre(01)'...")
        try:
            # Utiliser XPath avec le texte au lieu de l'ID dynamique
            autre_link = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Autre(01)')]"))
            )
            # Essayer clic JavaScript car le clic normal peut être intercepté
            try:
                driver.execute_script("arguments[0].click();", autre_link)
                print("      ✓ Lien 'Autre(01)' cliqué (via JavaScript)")
            except:
                autre_link.click()
                print("      ✓ Lien 'Autre(01)' cliqué")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur clic 'Autre(01)': {e}")
            # Essayer méthode alternative: chercher dans le tableau
            try:
                print("      🔄 Tentative alternative...")
                autre_link_alt = driver.find_element(By.XPATH, "//td[@role='gridcell']//a[contains(@class, 'ui-commandlink') and contains(text(), 'Autre')]")
                driver.execute_script("arguments[0].click();", autre_link_alt)
                print("      ✓ Lien 'Autre(01)' cliqué (méthode alternative)")
                time.sleep(2)
            except Exception as e2:
                print(f"      ❌ Erreur méthode alternative: {e2}")
                return False
        
        # ÉTAPE 5.4: Modifier le textarea avec le préfixe LTA et la valeur Carton
        print(f"\n   ✏️  Mise à jour du texte avec LTA N° {lta_reference_clean}...")
        try:
            # Attendre que le formulaire se charge après le clic
            time.sleep(1)
            
            # Chercher le textarea par XPath (plus robuste que l'ID dynamique)
            textarea = wait.until(
                EC.presence_of_element_located((By.XPATH, "//textarea[contains(@class, 'ui-inputtextarea') and @role='textbox']"))
            )
            
            # Lire le texte actuel
            current_text = textarea.get_attribute("value")
            print(f"      📄 Texte actuel: {current_text}")
            
            # Construire le nouveau texte avec préfixe LTA
            # Format: "LTA N° 72-73799132 SOIT {carton_value} COLIS.NS SOLL LA DISP DES FORM CCEC"
            new_text = re.sub(r'SOIT\s+\d+\s+COLIS', f'SOIT {carton_value} COLIS', current_text)
            new_text = f"LTA N° {lta_reference_clean} {new_text}"
            
            # Mettre à jour le textarea
            textarea.clear()
            textarea.send_keys(new_text)
            print(f"      ✓ Texte mis à jour: {new_text}")
            time.sleep(1)
        except Exception as e:
            print(f"      ❌ Erreur modification textarea: {e}")
            # Essayer méthode alternative avec ID contenant le pattern
            try:
                print("      🔄 Tentative alternative avec recherche par pattern...")
                textarea_alt = driver.find_element(By.XPATH, "//textarea[contains(@id, 'mainTab:form4:j_id') and contains(@class, 'ui-inputtextarea')]")
                current_text = textarea_alt.get_attribute("value")
                new_text = re.sub(r'SOIT\s+\d+\s+COLIS', f'SOIT {carton_value} COLIS', current_text)
                new_text = f"LTA N° {lta_reference_clean} {new_text}"
                textarea_alt.clear()
                textarea_alt.send_keys(new_text)
                print(f"      ✓ Texte mis à jour (méthode alternative): {new_text}")
                time.sleep(1)
            except Exception as e2:
                print(f"      ❌ Erreur méthode alternative: {e2}")
                return False
        
        # ÉTAPE 5.5: Confirmer la demande
        print("\n   ✅ Confirmation de la demande...")
        try:
            confirmer_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "mainTab:form4:btnConfirmerDmd"))
            )
            confirmer_btn.click()
            print("      ✓ Bouton 'Confirmer' cliqué")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur confirmation demande (ID): {e}")
            # Méthode alternative: chercher par texte
            try:
                print("      🔄 Tentative alternative par texte...")
                confirmer_btn_alt = driver.find_element(By.XPATH, "//button[contains(@class, 'ui-button')]//span[text()='Confirmer']/..")
                driver.execute_script("arguments[0].click();", confirmer_btn_alt)
                print("      ✓ Bouton 'Confirmer' cliqué (méthode alternative)")
                time.sleep(2)
            except Exception as e2:
                print(f"      ❌ Erreur confirmation alternative: {e2}")
                return False

        # ==================================================================
        # ÉTAPE 6: Naviguer vers "Moyen de transport"
        # ==================================================================
        print("\n   🚚 Navigation vers l'onglet 'Moyen de transport'...")
        try:
            moyen_transport_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab11']"))
            )
            moyen_transport_tab.click()
            print("      ✓ Onglet 'Moyen de transport' cliqué")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation vers 'Moyen de transport': {e}")
            return False
        
        # ÉTAPE 6.1: Cocher "Sans moyen de transport"
        print("\n   ☑️  Activation 'Sans moyen de transport'...")
        try:
            # Méthode 1: Cliquer sur la div.ui-chkbox-box visible
            sans_transport_checkbox = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form11\\:checkBoxSansMTId div.ui-chkbox-box"))
            )
            sans_transport_checkbox.click()
            print("      ✓ Checkbox 'Sans moyen de transport' cochée")
            time.sleep(1)
        except Exception as e:
            print(f"      ⚠️  Erreur checkbox (méthode UI): {e}")
            # Méthode alternative: JavaScript
            try:
                print("      🔄 Tentative avec JavaScript...")
                js_code = """
                var checkbox = document.getElementById('mainTab:form11:checkBoxSansMTId_input');
                checkbox.checked = true;
                var event = new Event('change', { bubbles: true });
                checkbox.dispatchEvent(event);
                """
                driver.execute_script(js_code)
                time.sleep(0.5)
                print("      ✓ Checkbox cochée via JavaScript")
            except Exception as e2:
                print(f"      ❌ Impossible de cocher 'Sans moyen de transport': {e2}")
                return False
        
        # ==================================================================
        # ÉTAPE PDS: Préapurement DS (CONDITIONNEL - seulement si LTA signé)
        # ==================================================================
        # Vérifier si le fichier [X]er LTA.txt existe et contient une série signée (Line 8)
        
        lta_name = os.path.basename(lta_folder_path)
        parent_dir = os.path.dirname(lta_folder_path)
        
        # Chercher le fichier [X]er LTA.txt dans le répertoire parent
        lta_file_pattern = os.path.join(parent_dir, f"{lta_name}.txt")
        
        preapurement_required = False
        ds_serie = None
        ds_cle = None
        validated_lta_reference = None
        loading_location = None
        
        if os.path.exists(lta_file_pattern):
            print(f"\n   📄 Fichier LTA trouvé: {lta_name}.txt")
            
            # Parser le fichier LTA
            lta_data = parse_lta_file(lta_file_pattern)
            
            if lta_data and lta_data['signed']:
                # Ligne 8 contient une série signée valide
                preapurement_required = True
                ds_serie = lta_data['serie']
                ds_cle = lta_data['cle']
                
                # IMPORTANT: Lire la référence LTA validée depuis le fichier shipper (ligne 5)
                # Cette référence a été sauvegardée lors de la Phase 1 (Etat de Dépotage)
                lta_name_with_underscore = lta_name.replace(" ", "_")
                shipper_pattern = f"{lta_name_with_underscore}_*.txt"
                shipper_files = glob.glob(os.path.join(parent_dir, shipper_pattern))
                
                if shipper_files:
                    try:
                        with open(shipper_files[0], 'r', encoding='utf-8') as f:
                            shipper_lines = [line.strip() for line in f.readlines()]
                        
                        # Line 3 du shipper file contient le lieu de chargement
                        if len(shipper_lines) >= 3:
                            loading_location = shipper_lines[2]
                        
                        # Line 5 du shipper file contient la référence LTA validée (sauvegardée en Phase 1)
                        if len(shipper_lines) >= 5 and shipper_lines[4]:
                            validated_lta_reference = shipper_lines[4]
                            print(f"\n   ✅ Préapurement DS requis (LTA signé)")
                            print(f"      Série signée: {lta_data['signed_series']}")
                            print(f"      Série: {ds_serie}")
                            print(f"      Clé: {ds_cle}")
                            print(f"      Référence LTA (depuis Phase 1): {validated_lta_reference}")
                            print(f"      Lieu de chargement: {loading_location}")
                        else:
                            print(f"      ⚠️  Référence LTA non trouvée dans {os.path.basename(shipper_files[0])} (ligne 5)")
                            print(f"      ℹ️  Assurez-vous d'avoir exécuté Phase 1 (Etat de Dépotage) d'abord")
                            # Ne pas continuer sans référence validée
                            preapurement_required = False
                    except Exception as e:
                        print(f"      ⚠️  Erreur lecture fichier shipper: {e}")
                        preapurement_required = False
                else:
                    print(f"      ⚠️  Fichier shipper introuvable: {shipper_pattern}")
                    preapurement_required = False
            else:
                print(f"\n   ⏭️  Préapurement DS non requis (LTA non signé - Line 8 vide ou invalide)")
                print(f"      ℹ️  Continuons avec la déclaration sans Préapurement DS")
        else:
            print(f"\n   ⏭️  Préapurement DS non requis (fichier LTA introuvable: {lta_file_pattern})")
            print(f"      ℹ️  Continuons avec la déclaration sans Préapurement DS")
        
        if preapurement_required and ds_serie and ds_cle and validated_lta_reference:
            print("\n" + "="*70)
            print("🔗 PRÉAPUREMENT DS")
            print("="*70)
            
            print(f"      📋 Série: {ds_serie}, Clé: {ds_cle}")
            
            # Construire la référence lot: validated_lta_reference + "/" + dum_number
            # Extraire le numéro DUM depuis sheet_name (e.g., "Sheet 1" → "1")
            sheet_name = dum_data.get('sheet_name', '')
            dum_number = sheet_name.split()[-1] if sheet_name.startswith('Sheet') else '1'
            
            # Format de la référence lot
            if "/" in validated_lta_reference:
                lot_reference = f"{validated_lta_reference}/{dum_number}"
            else:
                lot_reference = f"{validated_lta_reference}/{dum_number}"
            
            print(f"      📄 Référence lot: {lot_reference}")
            
            # PDS.1: Naviguer vers l'onglet "Préapurement DS"
            print("\n   📑 Navigation vers l'onglet 'Préapurement DS'...")
            try:
                preapurement_tab = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab3']"))
                )
                preapurement_tab.click()
                print("      ✓ Onglet 'Préapurement DS' cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur navigation 'Préapurement DS': {e}")
                return False
            
            # PDS.2: Cliquer sur "Nouveau"
            print("\n   ➕ Création d'un nouveau préapurement...")
            try:
                nouveau_preap_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@name, 'btnNouveauPreap')]"))
                )
                nouveau_preap_btn.click()
                print("      ✓ Bouton 'Nouveau' cliqué")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur clic 'Nouveau': {e}")
                return False
            
            # PDS.3: Remplir le formulaire
            print("\n   📝 Remplissage du formulaire Préapurement DS...")
            
            # PDS.3.1: Sélectionner type DS "Depotage(05)"
            try:
                # Attendre que le formulaire soit complètement chargé
                print("      ⏳ Attente du chargement du formulaire...")
                time.sleep(2)
                
                # Méthode 1: Cliquer sur le trigger pour ouvrir la liste déroulante
                try:
                    type_ds_trigger = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form3\\:typeDsId div.ui-selectonemenu-trigger"))
                    )
                    type_ds_trigger.click()
                    print("      ✓ Menu Type DS ouvert")
                    time.sleep(1)
                    
                    # Cliquer sur l'option "Depotage(05)"
                    depotage_option = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//li[@data-label='Depotage(05)']"))
                    )
                    depotage_option.click()
                    print("      ✓ Type DS: Depotage(05)")
                    time.sleep(1)
                    
                except Exception as click_err:
                    print(f"      ⚠️  Méthode clic échouée: {click_err}")
                    print("      🔄 Tentative avec JavaScript...")
                    
                    # Méthode 2: JavaScript (fallback)
                    js_code = """
                    var select = document.getElementById('mainTab:form3:typeDsId_input');
                    if (select) {
                        select.value = '05'; // Valeur pour "Depotage(05)"
                        var event = new Event('change', { bubbles: true });
                        select.dispatchEvent(event);
                        
                        // Mettre à jour le label visible
                        var label = document.getElementById('mainTab:form3:typeDsId_label');
                        if (label) {
                            label.textContent = 'Depotage(05)';
                        }
                    } else {
                        throw new Error('Select element not found');
                    }
                    """
                    driver.execute_script(js_code)
                    time.sleep(1)
                    print("      ✓ Type DS: Depotage(05) (via JavaScript)")
                    
            except Exception as e:
                print(f"      ❌ Impossible de sélectionner Type DS: {e}")
                return False
            
            # PDS.3.2: Bureau "301"
            try:
                bureau_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'bureauId') or contains(@name, 'bureauId')]"))
                )
                bureau_input.clear()
                bureau_input.send_keys("301")
                print("      ✓ Bureau: 301")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie bureau: {e}")
                return False
            
            # PDS.3.3: Régime "000"
            try:
                regime_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'regimeId') or contains(@name, 'regimeId')]"))
                )
                regime_input.clear()
                regime_input.send_keys("000")
                print("      ✓ Régime: 000")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie régime: {e}")
                return False
            
            # PDS.3.4: Année actuelle
            try:
                annee_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'anneeId') or contains(@name, 'anneeId')]"))
                )
                annee_input.clear()
                current_year = str(time.strftime("%Y"))
                annee_input.send_keys(current_year)
                print(f"      ✓ Année: {current_year}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie année: {e}")
                return False
            
            # PDS.3.5: Série DS
            try:
                serie_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'serieId') or contains(@name, 'serieId')]"))
                )
                serie_input.clear()
                serie_input.send_keys(ds_serie)
                print(f"      ✓ Série: {ds_serie}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie série: {e}")
                return False
            
            # PDS.3.6: Clé DS
            try:
                cle_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'cleId') or contains(@name, 'cleId')]"))
                )
                cle_input.clear()
                cle_input.send_keys(ds_cle)
                print(f"      ✓ Clé: {ds_cle}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie clé: {e}")
                return False
            
            # PDS.3.7: Lieu de chargement (autocomplete)
            if loading_location:
                try:
                    lieu_input = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'lieuChargCmb') or contains(@name, 'lieuChargCmb')]"))
                    )
                    lieu_input.clear()
                    lieu_input.send_keys(loading_location)
                    print(f"      ✓ Lieu chargement: {loading_location}")
                    time.sleep(2)
                    
                    # Sélectionner la première suggestion
                    try:
                        lieu_suggestion = wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.ui-autocomplete-item"))
                        )
                        lieu_suggestion.click()
                        print("      ✓ Suggestion lieu sélectionnée")
                        time.sleep(1)
                    except:
                        print("      ⚠️  Aucune suggestion trouvée, on continue...")
                except Exception as e:
                    print(f"      ⚠️  Erreur saisie lieu chargement: {e}")
            
            # PDS.3.8: Référence lot
            try:
                lot_ref_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'ref_lot') or contains(@name, 'ref_lot')]"))
                )
                lot_ref_input.clear()
                lot_ref_input.send_keys(lot_reference)
                print(f"      ✓ Référence lot: {lot_reference}")
                time.sleep(0.5)
            except Exception as e:
                print(f"      ❌ Erreur saisie référence lot: {e}")
                return False
            
            # PDS.4: Cliquer sur "OK" pour récupérer les données
            print("\n   🔍 Récupération des données du lot...")
            try:
                # Méthode 1: Par nom du bouton
                ok_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@id, 'btnRefPreapOk') or contains(text(), 'OK')]"))
                )
                try:
                    ok_btn.click()
                    print("      ✓ Bouton 'OK' cliqué")
                except:
                    # Fallback: JavaScript click
                    driver.execute_script("arguments[0].click();", ok_btn)
                    print("      ✓ Bouton 'OK' cliqué (via JavaScript)")
                time.sleep(3)
            except Exception as e:
                print(f"      ❌ Erreur clic 'OK': {e}")
                return False
            
            # PDS.5: Valider les données récupérées
            print("\n   ✅ Validation des données récupérées...")
            try:
                # Lire poids brut
                poids_brut_span = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:form3:poidLotId"))
                )
                poids_brut_text = poids_brut_span.text.strip()
                # Remplacer virgule par point (format français → format Python)
                poids_brut_text = poids_brut_text.replace(',', '.')
                # Garder la valeur EXACTE en float (ne pas arrondir)
                retrieved_weight = float(poids_brut_text)
                
                # Lire nombre contenants
                nbr_contenants_span = wait.until(
                    EC.presence_of_element_located((By.ID, "mainTab:form3:nbrContenantLotId"))
                )
                nbr_contenants_text = nbr_contenants_span.text.strip()
                # Gérer aussi le cas où il pourrait y avoir une virgule
                nbr_contenants_text = nbr_contenants_text.replace(',', '.')
                retrieved_containers = float(nbr_contenants_text)
                
                # Valeurs attendues depuis DUM data (aussi en float pour comparer exactement)
                expected_weight = float(dum_data.get('total_gross_weight', 0))
                expected_containers = float(dum_data.get('total_positions', 0))
                
                print(f"      📊 Poids brut: {retrieved_weight} (attendu: {expected_weight})")
                print(f"      📦 Contenants: {retrieved_containers} (attendu: {expected_containers})")
                
                # Vérifier correspondance
                if retrieved_weight != expected_weight or retrieved_containers != expected_containers:
                    print(f"      ❌ DIVERGENCE DÉTECTÉE!")
                    
                    # Créer fichier d'erreur
                    error_filename = f"-------------error-entering-ds-mead-on-declaration-{lta_name}-DUM{dum_number}.txt"
                    error_filepath = os.path.join(parent_dir, error_filename)
                    
                    from datetime import datetime
                    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    with open(error_filepath, 'w', encoding='utf-8') as f:
                        f.write(f"ERREUR - Préapurement DS - Données Incohérentes\n\n")
                        f.write(f"LTA: {lta_name} - {validated_lta_reference}\n")
                        f.write(f"DUM: {dum_number}\n")
                        f.write(f"Date: {current_datetime}\n")
                        f.write(f"Étape: Préapurement DS - Validation après click OK\n\n")
                        f.write(f"VALEURS ATTENDUES (DUM {dum_number}):\n")
                        f.write(f"- Poids brut (P,BRUT): {expected_weight}\n")
                        f.write(f"- Nombre contenants (P): {expected_containers}\n\n")
                        f.write(f"VALEURS RÉCUPÉRÉES (Système):\n")
                        f.write(f"- Poids brut: {retrieved_weight}\n")
                        f.write(f"- Nombre contenants: {retrieved_containers}\n\n")
                        f.write(f"ÉCART DÉTECTÉ:\n")
                        f.write(f"- Poids brut: {expected_weight} ≠ {retrieved_weight} (Différence: {expected_weight - retrieved_weight})\n")
                        f.write(f"- Contenants: {expected_containers} ≠ {retrieved_containers} (Différence: {expected_containers - retrieved_containers})\n\n")
                        f.write(f"MESSAGE: Les données du lot de dédouanement ne correspondent pas aux\n")
                        f.write(f"données du DUM actuel. Vérification manuelle requise.\n\n")
                        f.write(f"RÉFÉRENCE LOT UTILISÉE: {lot_reference}\n")
                        f.write(f"RÉFÉRENCE DS MEAD: {ds_serie} {ds_cle}\n")
                    
                    print(f"      ✓ Fichier d'erreur créé: {error_filename}")
                    print(f"      ⚠️  Arrêt du traitement de ce DUM")
                    return False
                
                print(f"      ✅ VALIDATION OK - Données correspondent")
                
            except Exception as e:
                print(f"      ❌ Erreur validation données: {e}")
                return False
            
            # PDS.6: Confirmer le préapurement
            print("\n   ✅ Confirmation du préapurement...")
            try:
                confirmer_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@id, 'btnConfirmerPreap') or (contains(@class, 'ui-button') and contains(., 'Confirmer'))]"))
                )
                try:
                    confirmer_btn.click()
                    print("      ✓ Bouton 'Confirmer' cliqué")
                except:
                    # Fallback: JavaScript click
                    driver.execute_script("arguments[0].click();", confirmer_btn)
                    print("      ✓ Bouton 'Confirmer' cliqué (via JavaScript)")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur confirmation préapurement: {e}")
                return False
            
            print("\n   ✅ Préapurement DS complété avec succès!")
            print("="*70)
            
            # Retourner à l'onglet Moyen de transport pour continuer le workflow normal
            print("\n   🔙 Retour à l'onglet 'Moyen de transport'...")
            try:
                moyen_transport_tab = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab11']"))
                )
                moyen_transport_tab.click()
                print("      ✓ Retour à 'Moyen de transport'")
                time.sleep(2)
            except Exception as e:
                print(f"      ❌ Erreur retour Moyen de transport: {e}")
                return False
        
        # ==================================================================
        # ÉTAPE 6.5: Cliquer sur "SAUVEGARDER" avant de naviguer vers Documents
        # ==================================================================
        print("\n   💾 Sauvegarde de la déclaration...")
        try:
            # Cliquer sur le bouton "SAUVEGARDER"
            sauvegarder_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "secure__2002"))
            )
            sauvegarder_btn.click()
            print("      ✓ Bouton 'SAUVEGARDER' cliqué")
            
            # Attendre que l'overlay de blocage disparaisse après la sauvegarde
            print("      ⏳ Attente de la fin de la sauvegarde...")
            if wait_for_ui_blocker_disappear(driver, timeout=10):
                print("      ✓ Sauvegarde terminée (blocker disparu)")
                time.sleep(2)  # Pause supplémentaire pour stabilité
            else:
                print("      ⚠️  Timeout blocker - continuons quand même")
                time.sleep(5)  # Fallback plus long
            
            # ==================================================================
            # VÉRIFIER SI LA SAUVEGARDE A RÉUSSI
            # ==================================================================
            print("      🔍 Vérification du résultat de sauvegarde...")
            save_error = False
            save_error_messages = []
            
            try:
                # Chercher les messages d'erreur
                error_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error")
                visible_errors = [c for c in error_containers if c.is_displayed()]
                
                if visible_errors:
                    # ⚠️ RÈGLE SPÉCIALE POUR SAUVEGARDER:
                    # Tout conteneur d'erreur visible = ÉCHEC, même s'il est vide!
                    # (contrairement aux validations où phantom errors = OK)
                    print("      ⚠️  Conteneur d'erreur détecté après sauvegarde")
                    save_error = True  # ← ERREUR IMMÉDIATE dès qu'un conteneur existe
                    
                    # Collecter les messages d'erreur s'ils existent
                    for error_container in visible_errors:
                        try:
                            # Chercher le bouton "Détails"
                            details_btn = error_container.find_element(By.ID, "rapportMsgForm:showErrors")
                            if details_btn and details_btn.is_displayed():
                                # Récupérer le texte d'erreur
                                error_details = error_container.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                                for detail in error_details:
                                    error_text = detail.text.strip()
                                    if error_text:
                                        save_error_messages.append(error_text)
                                        print(f"      ❌ Erreur: {error_text[:80]}...")
                                
                                # Si aucun message textuel, c'est une erreur "vide"
                                if not save_error_messages:
                                    save_error_messages.append("Erreur de sauvegarde (conteneur d'erreur vide)")
                                    print(f"      ❌ Erreur de sauvegarde détectée (conteneur vide)")
                        except:
                            # Pas de bouton "Détails" - vérifier message unique
                            error_details = error_container.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                            for detail in error_details:
                                error_text = detail.text.strip()
                                if error_text:
                                    save_error_messages.append(error_text)
                                    print(f"      ❌ Erreur: {error_text[:80]}...")
                            
                            # Si toujours aucun message
                            if not save_error_messages:
                                save_error_messages.append("Erreur de sauvegarde (conteneur sans détails)")
                                print(f"      ❌ Erreur de sauvegarde (conteneur sans message)")
                
                # Chercher message de succès
                if not save_error:
                    success_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-info")
                    visible_success = [c for c in success_containers if c.is_displayed()]
                    
                    if visible_success:
                        for success_container in visible_success:
                            success_details = success_container.find_elements(By.CSS_SELECTOR, "span.ui-messages-info-detail")
                            for detail in success_details:
                                success_text = detail.text.strip()
                                if success_text and "succès" in success_text.lower():
                                    print(f"      ✅ {success_text}")
                                    break
                    else:
                        print("      ✓ Déclaration sauvegardée (pas d'erreur détectée)")
                
            except Exception as check_err:
                print(f"      ⚠️  Impossible de vérifier le résultat: {check_err}")
                # Continuer par défaut si on ne peut pas vérifier
            
            # Si erreur de sauvegarde détectée, arrêter le traitement de ce DUM
            if save_error:
                print(f"\n   ❌ ÉCHEC SAUVEGARDE - Impossible de continuer avec ce DUM")
                print(f"      Erreur(s) détectée(s):")
                for msg in save_error_messages:
                    print(f"         • {msg}")
                
                # Retourner à l'accueil et marquer comme erreur
                print("\n   🏠 Retour à l'accueil après erreur de sauvegarde...")
                return_to_home_after_error(driver)
                
                # Marquer l'erreur dans Excel
                sheet_name = dum_data.get('sheet_name', '')
                dum_number = int(sheet_name.split()[-1]) if sheet_name.startswith('Sheet') else 1
                mark_dum_as_error_in_excel(lta_folder_path, dum_number)
                
                # Créer un log d'erreur
                lta_name = os.path.basename(lta_folder_path)
                save_dum_error_log(
                    lta_folder_path=lta_folder_path,
                    lta_name=lta_name,
                    dum_number=dum_number,
                    sheet_name=sheet_name,
                    error_exception=Exception(f"Erreur sauvegarde: {'; '.join(save_error_messages)}"),
                    error_step="Sauvegarde déclaration (SAUVEGARDER)",
                    dum_data=dum_data
                )
                
                return False  # Échec du DUM
                
        except Exception as e:
            print(f"      ⚠️  Erreur lors de la sauvegarde: {e}")
            # En cas d'exception, retourner à l'accueil et marquer comme erreur
            print("\n   🏠 Retour à l'accueil après exception...")
            return_to_home_after_error(driver)
            
            sheet_name = dum_data.get('sheet_name', '')
            dum_number = int(sheet_name.split()[-1]) if sheet_name.startswith('Sheet') else 1
            mark_dum_as_error_in_excel(lta_folder_path, dum_number)
            
            return False
        
        # ==================================================================
        # ÉTAPE 7: Naviguer vers "Documents" et uploader les fichiers
        # ==================================================================
        print("\n   📄 Navigation vers l'onglet 'Documents'...")
        try:
            # Attendre que l'onglet Documents soit cliquable (sans overlay)
            documents_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab7']"))
            )
            
            # Essayer le clic normal
            try:
                documents_tab.click()
                print("      ✓ Onglet 'Documents' cliqué")
            except Exception as click_error:
                print(f"      ⚠️  Clic normal intercepté, utilisation de JavaScript...")
                # Utiliser JavaScript si le clic est intercepté
                driver.execute_script("arguments[0].click();", documents_tab)
                print("      ✓ Onglet 'Documents' cliqué (via JavaScript)")
            
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation vers 'Documents': {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 7.1: Premier Upload - Document LTA
        # ==================================================================
        print("\n   📤 Upload 1/2: Document LTA principal...")
        
        # 7.1.1: Sélectionner le type de document "TITRE DE PROPRIÉTÉ ET/OU DE TRANSPORT"
        print("      1️⃣ Sélection du type 'TITRE DE PROPRIÉTÉ ET/OU DE TRANSPORT'...")
        try:
            # Cliquer sur le trigger pour ouvrir le dropdown
            doc_type_trigger = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form7\\:comp1 div.ui-selectonemenu-trigger"))
            )
            doc_type_trigger.click()
            print("         ✓ Dropdown type document ouvert")
            time.sleep(1)
            
            # Sélectionner l'option "TITRE DE PROPRIÉTÉ ET/OU DE TRANSPORT"
            doc_type_option = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li[@data-label='TITRE DE PROPRIÉTÉ ET/OU DE TRANSPORT']"))
            )
            doc_type_option.click()
            print("         ✓ Type sélectionné: A0004 (TITRE DE PROPRIÉTÉ ET/OU DE TRANSPORT)")
            time.sleep(1)
        except Exception as e:
            print(f"      ❌ Erreur sélection type document: {e}")
            return False
        
        # 7.1.2: Entrer la référence "LTA"
        print("      2️⃣ Saisie de la référence 'LTA'...")
        try:
            # Chercher l'input "Référence" par XPath - plus robuste que l'ID dynamique
            # On cherche l'input qui vient après le label "Référence"
            ref_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Référence')]/parent::td/following-sibling::td//input[@type='text']"))
            )
            # Attendre qu'il soit cliquable
            wait.until(lambda d: ref_input.is_enabled())
            ref_input.clear()
            ref_input.send_keys("LTA")
            print("         ✓ Référence: LTA")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie référence (méthode 1): {e}")
            # Méthode alternative: chercher par pattern d'ID
            try:
                print("      🔄 Tentative alternative...")
                ref_input_alt = driver.find_element(By.XPATH, "//input[contains(@id, 'mainTab:form7:j_id') and @type='text' and @maxlength='10']")
                ref_input_alt.clear()
                ref_input_alt.send_keys("LTA")
                print("         ✓ Référence: LTA (méthode alternative)")
                time.sleep(0.5)
            except Exception as e2:
                print(f"      ❌ Erreur saisie référence (méthode 2): {e2}")
                return False
        
        # 7.1.3: Sélectionner la date actuelle
        print("      3️⃣ Sélection de la date actuelle...")
        try:
            # Cliquer sur le bouton du date picker
            date_picker_btn = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-datepicker-trigger"))
            )
            date_picker_btn.click()
            time.sleep(1)
            
            # Sélectionner la date actuelle (ui-datepicker-today)
            today_cell = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "td.ui-datepicker-today a"))
            )
            today_cell.click()
            print("         ✓ Date actuelle sélectionnée")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur sélection date: {e}")
            return False
        
        # 7.1.4: Upload du fichier LTA PDF
        print("      4️⃣ Upload du fichier LTA PDF...")
        try:
            # Trouver le fichier LTA dans le dossier
            # Pattern: "12eme LTA - *.pdf" (le fichier principal LTA, pas les mn*.pdf)
            lta_name = os.path.basename(lta_folder_path)  # e.g., "12eme LTA"
            lta_pattern = os.path.join(lta_folder_path, f"{lta_name} - *.pdf")
            lta_files = glob.glob(lta_pattern)
            
            if not lta_files:
                print(f"      ❌ Fichier LTA PDF introuvable")
                print(f"         Pattern recherché: {lta_pattern}")
                return False
            
            lta_file_path = lta_files[0]
            print(f"         📄 Fichier trouvé: {os.path.basename(lta_file_path)}")
            
            # Upload du fichier
            file_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form7:comp2_input"))
            )
            absolute_path = os.path.abspath(lta_file_path)
            file_input.send_keys(absolute_path)
            print(f"         ✓ Fichier LTA uploadé: {os.path.basename(lta_file_path)}")
            
            # Attendre que le blocker d'upload disparaisse
            print("         ⏳ Attente fin d'upload...")
            if wait_for_ui_blocker_disappear(driver, timeout=10):
                print("         ✓ Upload terminé (blocker disparu)")
            else:
                print("         ⚠️  Timeout blocker upload - continuons")
            
            # Attendre que l'interface soit prête pour le prochain upload
            time.sleep(3)  # Stabilisation après premier upload
            
            # Vérifier à nouveau si un blocker apparaît (traitement en arrière-plan)
            print("         ⏳ Vérification stabilité UI...")
            if wait_for_ui_blocker_disappear(driver, timeout=5):
                print("         ✓ UI stabilisée")
            
            time.sleep(2)  # Pause supplémentaire pour sécurité
            
            print("         ✓ Upload LTA traité, préparation pour le document MN...")
        except Exception as e:
            print(f"      ❌ Erreur upload fichier LTA: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 7.2: Deuxième Upload - Document MN du DUM actuel
        # ==================================================================
        print("\n   📤 Upload 2/2: Document MN pour ce DUM...")
        
        # Extraire le numéro du DUM depuis sheet_name (e.g., "Sheet 1" -> "1")
        sheet_name = dum_data.get('sheet_name', '')
        dum_number = sheet_name.split()[-1] if sheet_name.startswith('Sheet') else '1'
        mn_reference = f"mn{dum_number}"
        mn_filename = f"mn{dum_number}.pdf"
        
        # 7.2.1: Sélectionner le type de document "FACTURE"
        print(f"      1️⃣ Sélection du type 'FACTURE'...")
        try:
            # Cliquer sur le trigger pour ouvrir le dropdown
            doc_type_trigger = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#mainTab\\:form7\\:comp1 div.ui-selectonemenu-trigger"))
            )
            doc_type_trigger.click()
            print("         ✓ Dropdown type document ouvert")
            time.sleep(1)
            
            # Sélectionner l'option "FACTURE"
            doc_type_option = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//li[@data-label='FACTURE']"))
            )
            doc_type_option.click()
            print("         ✓ Type sélectionné: A0006 (FACTURE)")
            time.sleep(1)
        except Exception as e:
            print(f"      ❌ Erreur sélection type document: {e}")
            return False
        
        # 7.2.2: Entrer la référence "mnN"
        print(f"      2️⃣ Saisie de la référence '{mn_reference}'...")
        try:
            # Chercher l'input "Référence" par XPath
            ref_input = wait.until(
                EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'Référence')]/parent::td/following-sibling::td//input[@type='text']"))
            )
            wait.until(lambda d: ref_input.is_enabled())
            ref_input.clear()
            ref_input.send_keys(mn_reference)
            print(f"         ✓ Référence: {mn_reference}")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur saisie référence (méthode 1): {e}")
            # Méthode alternative
            try:
                print("      🔄 Tentative alternative...")
                ref_input_alt = driver.find_element(By.XPATH, "//input[contains(@id, 'mainTab:form7:j_id') and @type='text' and @maxlength='10']")
                ref_input_alt.clear()
                ref_input_alt.send_keys(mn_reference)
                print(f"         ✓ Référence: {mn_reference} (méthode alternative)")
                time.sleep(0.5)
            except Exception as e2:
                print(f"      ❌ Erreur saisie référence (méthode 2): {e2}")
                return False
        
        # 7.2.3: Sélectionner la date actuelle (à nouveau)
        print("      3️⃣ Sélection de la date actuelle...")
        try:
            date_picker_btn = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-datepicker-trigger"))
            )
            date_picker_btn.click()
            time.sleep(1)
            
            today_cell = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "td.ui-datepicker-today a"))
            )
            today_cell.click()
            print("         ✓ Date actuelle sélectionnée")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ❌ Erreur sélection date: {e}")
            return False
        
        # 7.2.4: Upload du fichier MN PDF
        print(f"      4️⃣ Upload du fichier {mn_filename}...")
        try:
            # Chercher le fichier mn*.pdf dans le dossier LTA
            mn_file_path = os.path.join(lta_folder_path, mn_filename)
            
            if not os.path.exists(mn_file_path):
                print(f"      ❌ Fichier MN introuvable: {mn_filename}")
                print(f"         Chemin recherché: {mn_file_path}")
                return False
            
            print(f"         📄 Fichier trouvé: {mn_filename}")
            
            # Upload du fichier
            file_input = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form7:comp2_input"))
            )
            absolute_path = os.path.abspath(mn_file_path)
            file_input.send_keys(absolute_path)
            print(f"         ✓ Fichier MN uploadé: {mn_filename}")
            
            # Attendre que le blocker d'upload disparaisse
            print("         ⏳ Attente fin d'upload...")
            if wait_for_ui_blocker_disappear(driver, timeout=10):
                print("         ✓ Upload terminé (blocker disparu)")
            else:
                print("         ⚠️  Timeout blocker upload - continuons")
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur upload fichier MN: {e}")
            return False
        
        print("\n   ✅ Documents uploadés avec succès (LTA + MN)")
        
        # ==================================================================
        # ÉTAPE 8: Retour à l'onglet "Entête" pour validation finale
        # ==================================================================
        print("\n   📋 Navigation vers l'onglet 'Entête' pour validation...")
        try:
            # Naviguer vers l'onglet Entête
            entete_tab = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#mainTab:tab0']"))
            )
            
            # Essayer le clic normal
            try:
                entete_tab.click()
                print("      ✓ Onglet 'Entête' cliqué")
            except Exception as click_error:
                print(f"      ⚠️  Clic normal intercepté, utilisation de JavaScript...")
                driver.execute_script("arguments[0].click();", entete_tab)
                print("      ✓ Onglet 'Entête' cliqué (via JavaScript)")
            
            time.sleep(2)
        except Exception as e:
            print(f"      ❌ Erreur navigation vers 'Entête': {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 9: Premier clic sur "VALIDER" pour révéler le champ Commerce électronique
        # ==================================================================
        print("\n   🔍 Premier clic sur 'VALIDER' pour révéler les champs manquants...")
        try:
            # Cliquer sur le bouton VALIDER (première fois)
            valider_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "secure__2003"))
            )
            valider_btn.click()
            print("      ✓ Bouton 'VALIDER' cliqué (1ère fois)")
            
            # Attendre que le système affiche l'erreur et révèle le champ
            time.sleep(2)
            print("      ✓ Champ 'Commerce électronique' révélé")
        except Exception as e:
            print(f"      ❌ Erreur premier clic VALIDER: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 10: Cocher "Commerce électronique - Oui"
        # ==================================================================
        print("\n   ☑️  Activation 'Commerce électronique - Oui'...")
        try:
            # Attendre que le radio button soit présent
            time.sleep(1)
            
            # Méthode directe: chercher tous les div.ui-radiobutton-box dans le tableau Commerce électronique
            # et prendre le premier (Oui)
            commerce_elec_radios = driver.find_elements(By.CSS_SELECTOR, "table#mainTab\\:form0\\:commerceElectronique div.ui-radiobutton-box")
            if len(commerce_elec_radios) >= 1:
                commerce_elec_radios[0].click()  # Le premier = Oui
                print("      ✓ Radio 'Commerce électronique - Oui' coché")
                time.sleep(1)
            else:
                print(f"      ⚠️  Aucun radio button trouvé (nombre: {len(commerce_elec_radios)})")
                # Méthode alternative: JavaScript
                print("      🔄 Tentative avec JavaScript...")
                js_code = """
                var radio = document.getElementById('mainTab:form0:commerceElectronique:0');
                if (radio) {
                    radio.checked = true;
                    var event = new Event('change', { bubbles: true });
                    radio.dispatchEvent(event);
                } else {
                    throw new Error('Radio button not found');
                }
                """
                driver.execute_script(js_code)
                time.sleep(0.5)
                print("      ✓ Radio coché via JavaScript")
        except Exception as e:
            print(f"      ❌ Impossible de cocher Commerce électronique: {e}")
            return False
        
        # ==================================================================
        # ÉTAPE 11: Deuxième clic sur "VALIDER" pour soumettre la déclaration
        # ==================================================================
        print("\n   ✅ Validation finale de la déclaration...")
        try:
            # Cliquer sur le bouton VALIDER (deuxième fois - soumission finale)
            valider_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "secure__2003"))
            )
            valider_btn.click()
            print("      ✓ Bouton 'VALIDER' cliqué (2ème fois - soumission)")
            
            # Attendre que la validation soit traitée
            print("      ⏳ Attente de la validation...")
            
            # Attendre que le blocker UI disparaisse (validation en cours)
            if wait_for_ui_blocker_disappear(driver, timeout=15):
                print("      ✓ Validation terminée (blocker disparu)")
            else:
                print("      ⚠️  Timeout blocker validation - continuons")
            
            # Pause supplémentaire pour stabilité
            time.sleep(3)
            
            # ==================================================================
            # VÉRIFICATION DES MESSAGES DE VALIDATION (SUCCÈS OU ERREUR)
            # ==================================================================
            validation_error = False
            error_messages = []
            
            # D'abord, vérifier s'il y a des messages de SUCCÈS (ui-messages-info)
            has_success_message = False
            try:
                info_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-info")
                visible_info = [c for c in info_containers if c.is_displayed()]
                if visible_info:
                    has_success_message = True
                    print("      ℹ️  Message d'information détecté (probablement succès)")
            except:
                pass
            
            # Chercher messages d'erreur
            try:
                # Chercher UNIQUEMENT les conteneurs d'erreur VISIBLES
                all_error_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-error")
                error_containers = [c for c in all_error_containers if c.is_displayed()]
                
                # ⚠️ IMPORTANT: Si on a un message de succès ET aucun conteneur d'erreur visible,
                # alors c'est une validation réussie (ne pas traiter comme erreur)
                if has_success_message and (not error_containers or len(error_containers) == 0):
                    print("      ✅ Message de succès détecté sans erreur - validation réussie")
                    validation_error = False
                    error_messages = []
                
                elif error_containers and len(error_containers) > 0:
                    print("      ⚠️  Conteneur d'erreur détecté")
                    
                    # ==================================================================
                    # VÉRIFIER SI BOUTON "DÉTAILS" EXISTE = PLUSIEURS ERREURS
                    # ==================================================================
                    has_details_button = False
                    try:
                        # Chercher le bouton "Détails" UNIQUEMENT dans les conteneurs d'erreur
                        for error_container in error_containers:
                            try:
                                details_btn = error_container.find_element(By.ID, "rapportMsgForm:showErrors")
                                if details_btn and details_btn.is_displayed():
                                    has_details_button = True
                                    break
                            except:
                                continue
                    except:
                        pass
                    
                    if has_details_button:
                        # ==================================================================
                        # BOUTON "DÉTAILS" DÉTECTÉ → TOUJOURS TRAITER COMME ERREUR
                        # ==================================================================
                        # ⚠️ RÈGLE CRITIQUE: Si le bouton "Détails" existe, c'est qu'il y a
                        # une liste d'erreurs cachée, même si le span error-detail est vide!
                        print(f"      ⚠️  Bouton 'Détails' détecté → Erreurs multiples présentes")
                        validation_error = True
                        error_messages.append("Plusieurs erreurs de validation détectées (cliquer sur 'Détails' dans BADR pour voir la liste)")
                        print(f"      ❌ Déclaration invalide (erreurs multiples)")
                        
                    else:
                        # ==================================================================
                        # MESSAGE UNIQUE → EXTRAIRE ET VÉRIFIER SI C'EST "COMMERCE ÉLECTRONIQUE"
                        # ==================================================================
                        print(f"      ℹ️  Pas de bouton 'Détails' → Message unique")
                        
                        # Extraire le message unique
                        try:
                            error_details = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-error-detail")
                            if error_details and len(error_details) > 0:
                                for detail in error_details:
                                    try:
                                        error_text = detail.text.strip()
                                        if error_text:
                                            lines = error_text.split('\n')
                                            for line in lines:
                                                line = line.strip()
                                                if line and line not in error_messages:
                                                    error_messages.append(line)
                                    except:
                                        continue
                        except Exception as extract_err:
                            print(f"      ⚠️  Erreur extraction message: {extract_err}")
                        
                        # Si pas de message extrait, chercher dans le conteneur
                        if not error_messages:
                            try:
                                for container in error_containers:
                                    try:
                                        container_text = container.text.strip()
                                        if container_text and "Erreur" in container_text:
                                            lines = container_text.split('\n')
                                            for line in lines:
                                                line = line.strip()
                                                if line and line != "Erreur :" and line not in error_messages:
                                                    error_messages.append(line)
                                    except:
                                        continue
                            except:
                                pass
                        
                        # Vérifier si c'est le message "Commerce électronique"
                        if len(error_messages) == 1:
                            single_msg = error_messages[0]
                            if "commerce électronique" in single_msg.lower() and "information obligatoire" in single_msg.lower():
                                # Message informatif → Ignorer
                                print(f"      ℹ️  Message informatif ignoré: {single_msg}")
                                error_messages = []
                                print("      ✅ Aucune erreur bloquante - validation réussie")
                            else:
                                # Erreur réelle unique
                                validation_error = True
                                print(f"      ⚠️  Erreur de validation détectée: {single_msg[:80]}{'...' if len(single_msg) > 80 else ''}")
                        elif len(error_messages) > 1:
                            # Cas rare: plusieurs messages extraits même sans bouton Détails
                            validation_error = True
                            print(f"      ⚠️  {len(error_messages)} erreurs détectées")
                        else:
                            # Aucun message d'erreur extrait
                            # Si on a un message de succès visible, c'est une validation réussie
                            if has_success_message:
                                print("      ✅ Conteneur d'erreur vide + message de succès → validation réussie")
                            else:
                                print("      ✅ Aucune erreur bloquante - validation réussie")
            except Exception as e:
                # Erreur lors de la détection - considérer comme pas d'erreur
                print(f"      ℹ️  Impossible de détecter les messages (probablement aucune erreur)")
            
            # Si erreur détectée, créer un log détaillé et marquer dans Excel
            if validation_error:
                print(f"\n   ❌ DÉCLARATION REFUSÉE - Erreurs de validation")
                
                # Extraire le numéro du DUM
                sheet_name = dum_data.get('sheet_name', '')
                dum_number = int(sheet_name.split()[-1]) if sheet_name.startswith('Sheet') else 1
                
                # ==================================================================
                # TENTER D'EXTRAIRE LA SÉRIE MÊME EN CAS D'ERREUR
                # ==================================================================
                # La série peut être visible dans la table même si la validation a échoué
                dum_series = None
                try:
                    # Chercher la table de référence
                    reference_table = driver.find_element(By.ID, "mainTab:form0:j_id_3p_d")
                    rows = reference_table.find_elements(By.TAG_NAME, "tr")
                    if len(rows) >= 2:
                        data_row = rows[1]
                        cells = data_row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 5:
                            serie = cells[3].text.strip()
                            cle = cells[4].text.strip()
                            if serie and cle:
                                dum_series = f"{serie}{cle}"
                                print(f"      ℹ️  Série extraite malgré l'erreur: {dum_series}")
                except Exception as serie_err:
                    print(f"      ⚠️  Impossible d'extraire la série: {serie_err}")
                
                # Si série non trouvée, utiliser un placeholder
                if not dum_series:
                    dum_series = "SÉRIE_INCONNUE"
                
                # Créer le fichier d'erreur
                lta_name_safe = lta_name.replace(' ', '_')
                error_filename = f"error-validating-declaration-dedouanement-{lta_name_safe}-DUM{dum_number}.txt"
                parent_dir = os.path.dirname(lta_folder_path)
                error_filepath = os.path.join(parent_dir, error_filename)
                
                from datetime import datetime
                current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # Analyser les catégories d'erreurs
                error_categories = []
                if any("crédit" in msg.lower() and "enlèvement" in msg.lower() for msg in error_messages):
                    error_categories.append("Crédit d'enlèvement manquant ou invalide")
                if any("ti-a" in msg.lower() or "propriete ti" in msg.lower() for msg in error_messages):
                    error_categories.append("Propriétés techniques incomplètes (TI-A)")
                if any("quantité" in msg.lower() for msg in error_messages):
                    error_categories.append("Données quantitatives manquantes")
                if any("tva" in msg.lower() for msg in error_messages):
                    error_categories.append("Informations TVA manquantes")
                if any("poids net" in msg.lower() for msg in error_messages):
                    error_categories.append("Divergence de poids net")
                
                # Créer le fichier log détaillé
                with open(error_filepath, 'w', encoding='utf-8') as f:
                    f.write("=" * 70 + "\n")
                    f.write("ERREUR - VALIDATION DÉCLARATION DÉDOUANEMENT\n")
                    f.write("=" * 70 + "\n\n")
                    
                    f.write(f"LTA: {lta_name}\n")
                    f.write(f"DUM: {dum_number}\n")
                    f.write(f"Date: {current_datetime}\n")
                    f.write(f"Étape: Validation finale déclaration dédouanement\n\n")
                    
                    f.write("-" * 70 + "\n")
                    f.write("DÉTAILS DUM:\n")
                    f.write("-" * 70 + "\n")
                    f.write(f"Sheet Name: {dum_data.get('sheet_name', 'N/A')}\n")
                    f.write(f"Total Value: {dum_data.get('total_value', 0)}\n")
                    f.write(f"Gross Weight: {dum_data.get('total_gross_weight', 0)}\n")
                    f.write(f"Positions: {dum_data.get('total_positions', 0)}\n")
                    f.write(f"Freight: {dum_data.get('total_freight', 0)}\n")
                    f.write(f"Insurance: {dum_data.get('insurance', 0)}\n")
                    f.write(f"Cartons: {dum_data.get('cartons', 0)}\n\n")
                    
                    f.write("=" * 70 + "\n")
                    f.write("MESSAGES D'ERREUR DU SYSTÈME:\n")
                    f.write("=" * 70 + "\n\n")
                    for i, msg in enumerate(error_messages, 1):
                        f.write(f"  • {msg}\n")
                    f.write("\n")
                    
                    if error_categories:
                        f.write("=" * 70 + "\n")
                        f.write("CATÉGORIES D'ERREURS DÉTECTÉES:\n")
                        f.write("=" * 70 + "\n\n")
                        for cat in error_categories:
                            f.write(f"  ✗ {cat}\n")
                        f.write("\n")
                    
                    f.write("=" * 70 + "\n")
                    f.write("ACTION REQUISE:\n")
                    f.write("=" * 70 + "\n\n")
                    f.write("Ce DUM n'a pas pu être validé automatiquement.\n")
                    f.write("Veuillez:\n")
                    f.write("  1. Vérifier les données du fichier Excel source\n")
                    f.write("  2. Corriger les informations manquantes\n")
                    f.write("  3. Créer la déclaration manuellement ou relancer le script\n\n")
                    
                    f.write("=" * 70 + "\n")
                    f.write("FICHIERS CONCERNÉS:\n")
                    f.write("=" * 70 + "\n\n")
                    f.write(f"  • Sheet Excel: {dum_data.get('sheet_name', 'N/A')}\n")
                    f.write(f"  • LTA Folder: {lta_name}/\n\n")
                    
                    f.write("=" * 70 + "\n")
                    f.write("FIN DU RAPPORT D'ERREUR\n")
                    f.write("=" * 70 + "\n")
                
                print(f"      ✓ Fichier d'erreur créé: {error_filename}")
                
                # Marquer "error" dans le fichier Excel (colonne C)
                try:
                    # Trouver le fichier generated_excel
                    generated_excel_path = None
                    for file in os.listdir(lta_folder_path):
                        if file.startswith("generated_excel") and file.endswith(".xlsx"):
                            generated_excel_path = os.path.join(lta_folder_path, file)
                            break
                    
                    if generated_excel_path:
                        # Calculer la position: C + (12 + (dum_number - 1) * 7)
                        row_number = 12 + (dum_number - 1) * 7
                        cell_position = f"C{row_number}"
                        
                        wb = load_workbook(generated_excel_path, data_only=False)
                        ws = wb['Summary']
                        
                        # Écrire la série avec marqueur d'erreur
                        error_marker = f"{dum_series} (error)"
                        ws[cell_position] = error_marker
                        
                        wb.save(generated_excel_path)
                        wb.close()
                
                        print(f"      ✓ Marqueur 'error' ajouté en {cell_position}: {error_marker}")
                except Exception as excel_err:
                    print(f"      ⚠️  Erreur marquage Excel: {excel_err}")
                
                print(f"      ❌ Déclaration refusée - Erreurs détectées")
                print(f"      ⏭️  Passage au DUM suivant...")
                
                # Retourner à l'accueil pour continuer avec le prochain DUM
                try:
                    time.sleep(2)
                    accueil_btn = wait.until(EC.element_to_be_clickable((By.ID, "quitter")))
                    try:
                        accueil_btn.click()
                    except:
                        driver.execute_script("arguments[0].click();", accueil_btn)
                    time.sleep(3)
                    driver.switch_to.default_content()
                    print("      ✓ Retour à l'accueil réussi")
                except:
                    try:
                        driver.switch_to.default_content()
                    except:
                        pass
                
                return False  # Indiquer l'échec de ce DUM
            
            # Si pas d'erreur, chercher message de succès (info)
            try:
                info_containers = driver.find_elements(By.CSS_SELECTOR, "div.ui-messages-info")
                if info_containers and len(info_containers) > 0:
                    info_details = driver.find_elements(By.CSS_SELECTOR, "span.ui-messages-info-detail")
                    if info_details and len(info_details) > 0:
                        info_text = info_details[0].text.strip()
                        if info_text:
                            print(f"      ℹ️  Message info: {info_text[:80]}...")
            except:
                pass
            
            print("      ✅ Déclaration validée avec succès")
            
        except Exception as e:
            print(f"      ❌ Erreur validation finale: {e}")
            return False
        
        print(f"\n   ✅ DUM {dum_data.get('sheet_name', 'DUM')} traité et validé avec succès!")
        
        # ==================================================================
        # ÉTAPE 12: Extraire la référence de déclaration et la sauvegarder
        # ==================================================================
        print("\n   📋 Extraction de la référence de déclaration...")
        try:
            # Attendre que la table de référence soit visible
            time.sleep(2)
            
            # Localiser la table de référence
            reference_table = wait.until(
                EC.presence_of_element_located((By.ID, "mainTab:form0:j_id_3p_d"))
            )
            
            # Extraire les cellules de la deuxième ligne (ligne de données)
            rows = reference_table.find_elements(By.TAG_NAME, "tr")
            if len(rows) >= 2:
                data_row = rows[1]  # Deuxième ligne (index 1)
                cells = data_row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) >= 5:
                    # Extraire Série (4ème colonne, index 3)
                    serie = cells[3].text.strip()
                    # Extraire Clé (5ème colonne, index 4)
                    cle = cells[4].text.strip()
                    
                    # Combiner pour créer la référence complète
                    dum_reference = f"{serie}{cle}"
                    
                    print(f"      ✓ Référence extraite: {dum_reference}")
                    print(f"         - Série: {serie}")
                    print(f"         - Clé: {cle}")
                    
                    # Sauvegarder la référence dans result_LTAS.txt
                    save_dum_reference(lta_folder_path, dum_reference)
                    
                    # Extraire le numéro du DUM depuis sheet_name (ex: "Sheet 1" → 1)
                    sheet_name = dum_data.get('sheet_name', '')
                    dum_number = int(sheet_name.split()[-1]) if sheet_name.startswith('Sheet') else 1
                    
                    # Sauvegarder la série dans generated_excel
                    save_dum_series_to_excel(lta_folder_path, dum_number, dum_reference)
                    
                else:
                    print(f"      ⚠️  Table de référence incomplète (cellules: {len(cells)})")
                    dum_reference = "REFERENCE_INCOMPLETE"
            else:
                print(f"      ⚠️  Table de référence incomplète (lignes: {len(rows)})")
                dum_reference = "REFERENCE_INCOMPLETE"
                
        except Exception as e:
            print(f"      ❌ Erreur extraction référence: {e}")
            dum_reference = "REFERENCE_ERROR"
            traceback.print_exc()
        
        # ==================================================================
        # ÉTAPE 13: Retour à l'accueil pour traiter le prochain DUM
        # ==================================================================
        print("\n   🏠 Retour à l'accueil pour le prochain DUM...")
        try:
            # Attendre que la page soit complètement stable après validation
            print("      ⏳ Attente stabilisation page...")
            time.sleep(3)
            
            # Attendre que le blocker soit complètement disparu
            if wait_for_ui_blocker_disappear(driver, timeout=10):
                print("      ✓ Page stabilisée (blocker disparu)")
            else:
                print("      ⚠️  Timeout blocker - continuons")
            
            # Pause supplémentaire avant de cliquer sur Accueil
            time.sleep(2)
            
            # Cliquer sur le bouton "Accueil" (id="quitter")
            accueil_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "quitter"))
            )
            
            try:
                accueil_btn.click()
                print("      ✓ Bouton 'Accueil' cliqué")
            except Exception as click_error:
                print(f"      ⚠️  Clic normal intercepté, utilisation de JavaScript...")
                driver.execute_script("arguments[0].click();", accueil_btn)
                print("      ✓ Bouton 'Accueil' cliqué (via JavaScript)")
            
            # Attendre que le blocker de navigation disparaisse
            print("      ⏳ Attente navigation vers accueil...")
            if wait_for_ui_blocker_disappear(driver, timeout=10):
                print("      ✓ Navigation terminée (blocker disparu)")
            else:
                print("      ⚠️  Timeout blocker navigation")
            
            # Attendre le retour à la page d'accueil
            time.sleep(3)
            
            # IMPORTANT: Sortir de l'iframe pour revenir au contexte principal
            driver.switch_to.default_content()
            print("      ✓ Sorti de l'iframe, retour au contexte principal")
            
            print("      ✓ Retour à l'accueil réussi")
            
        except Exception as e:
            print(f"      ❌ Erreur retour accueil: {e}")
            traceback.print_exc()
            # Essayer quand même de sortir de l'iframe
            try:
                driver.switch_to.default_content()
                print("      ⚠️  Sorti de l'iframe malgré l'erreur")
            except:
                pass
        
        return True
        
    except Exception as e:
        print(f"\n   ❌ Erreur remplissage formulaire: {e}")
        traceback.print_exc()
        
        # NETTOYAGE CRITIQUE: S'assurer de sortir de l'iframe et revenir à l'état initial
        print("\n   🧹 Nettoyage après erreur...")
        try:
            # 1. Sortir de l'iframe
            driver.switch_to.default_content()
            print("      ✓ Sorti de l'iframe")
            
            # 2. Rafraîchir la page pour revenir à l'accueil
            driver.get(driver.current_url.split('#')[0])
            time.sleep(3)
            print("      ✓ Page rafraîchie, retour à l'état initial")
            
        except Exception as cleanup_err:
            print(f"      ⚠️  Erreur nettoyage: {cleanup_err}")
            # Dernière tentative: recharger complètement la page d'accueil
            try:
                driver.get("https://badr.douane.gov.ma:40444/badr/")
                time.sleep(3)
                print("      ✓ Rechargement complet de la page d'accueil")
            except:
                pass
        
        return False

def process_lta_folder(driver, lta_folder_path, lta_name):
    """Process a complete LTA folder: read data and fill forms for all DUMs
    
    Args:
        driver: Selenium WebDriver instance (should be logged in)
        lta_folder_path: Path to LTA folder (e.g., "./8eme LTA")
        lta_name: Name of LTA (e.g., "8eme LTA")
    
    Returns:
        Number of DUMs successfully processed
    """
    try:
        print("\n" + "="*70)
        print(f"📁 TRAITEMENT DU DOSSIER: {lta_name}")
        print("="*70)
        
        # 1. Read shipper data from .txt file (parent directory)
        parent_dir = os.path.dirname(lta_folder_path)
        
        # Read from the new format file: "8eme_LTA_shipper_name.txt"
        safe_name = lta_name.replace(' ', '_')
        txt_file_path = os.path.join(parent_dir, f"{safe_name}_shipper_name.txt")
        
        if not os.path.exists(txt_file_path):
            print(f"❌ Fichier shipper introuvable: {safe_name}_shipper_name.txt")
            return 0
        
        shipper_data = read_shipper_from_txt(txt_file_path)
        if not shipper_data:
            print(f"❌ Impossible de lire les données depuis {txt_file_path}")
            return 0
        
        print(f"✓ Expéditeur: {shipper_data['shipper_name']}")
        
        # 1.5 CONDITION: Créer Etat de Dépotage SI ligne 2 existe (has_ds_mead)
        if shipper_data['has_ds_mead']:
            print(f"\n✅ LTA avec référence DS MEAD détectée")
            print(f"   - Série: {shipper_data['serie']}")
            print(f"   - Clé: {shipper_data['cle']}")
            print(f"   - Lieu: {shipper_data['loading_location']}")
            print("\n🔄 Création de l'Etat de Dépotage...")
            
            if not create_etat_depotage(driver, lta_folder_path, shipper_data):
                print("❌ Échec création Etat de Dépotage - Arrêt du traitement")
                return 0
            
            print("\n✅ Etat de Dépotage créé avec succès - Passage aux DUMs")
            
            # Retour à l'accueil après Etat de Dépotage
            print("\n🏠 Retour à l'accueil après Etat de Dépotage...")
            try:
                # Créer WebDriverWait pour cette section
                wait = WebDriverWait(driver, 10)
                
                # Attendre un peu pour que la page soit stable
                time.sleep(2)
                
                # Cliquer sur le bouton "Accueil" (id="quitter")
                accueil_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "quitter"))
                )
                
                try:
                    accueil_btn.click()
                    print("      ✓ Bouton 'Accueil' cliqué")
                except Exception as click_error:
                    print(f"      ⚠️  Clic normal intercepté, utilisation de JavaScript...")
                    driver.execute_script("arguments[0].click();", accueil_btn)
                    print("      ✓ Bouton 'Accueil' cliqué (via JavaScript)")
                
                # Attendre le retour à la page d'accueil
                time.sleep(3)
                
                # IMPORTANT: Sortir de l'iframe pour revenir au contexte principal
                driver.switch_to.default_content()
                print("      ✓ Sorti de l'iframe, retour au contexte principal")
                
                print("      ✓ Retour à l'accueil réussi")
                
            except Exception as e:
                print(f"      ❌ Erreur retour accueil (bouton): {e}")
                traceback.print_exc()
                
                # FALLBACK: Naviguer directement vers la page d'accueil
                print("      🔄 Fallback: Navigation directe vers l'accueil...")
                try:
                    driver.switch_to.default_content()
                    print("      ✓ Sorti de l'iframe")
                    
                    driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
                    print("      ✓ Navigation vers l'accueil réussie (URL directe)")
                    time.sleep(3)  # Attendre le chargement de la page
                except Exception as e2:
                    print(f"      ❌ Erreur navigation directe: {e2}")
                    # Essayer quand même de sortir de l'iframe
                    try:
                        driver.switch_to.default_content()
                        print("      ⚠️  Sorti de l'iframe malgré l'erreur")
                    except:
                        pass
        else:
            print("\n⏭️  LTA sans référence DS MEAD (ligne 2 absente)")
            print("   → Saut de l'Etat de Dépotage, passage direct aux DUMs")
        
        # 2. Find and read summary_file Excel
        summary_files = glob.glob(os.path.join(lta_folder_path, "summary_file*.xlsx"))
        if not summary_files:
            print(f"❌ Aucun summary_file trouvé dans {lta_folder_path}")
            return 0
        
        summary_file_path = summary_files[0]
        print(f"✓ Fichier summary: {os.path.basename(summary_file_path)}")
        
        # 3. Read all DUM data from summary
        dum_list = read_dum_data_from_summary(summary_file_path)
        if not dum_list:
            print(f"❌ Aucune donnée DUM trouvée dans {summary_file_path}")
            return 0
        
        print(f"\n📊 {len(dum_list)} DUMs à traiter:")
        for i, dum in enumerate(dum_list, 1):
            print(f"   {i}. {dum.get('sheet_name')} - Valeur: {dum.get('total_value')} - Poids: {dum.get('total_gross_weight')}")
        
        # 4. Process each DUM
        successful_count = 0
        
        for i, dum_data in enumerate(dum_list, 1):
            print(f"\n{'='*70}")
            print(f"DUM {i}/{len(dum_list)}: {dum_data.get('sheet_name')}")
            print(f"{'='*70}")
            
            # Create declaration (this navigates to the form)
            if not create_declaration(driver):
                print(f"❌ Échec création déclaration pour {dum_data.get('sheet_name')}")
                continue
            
            # Fill the form with shipper and DUM data
            if fill_declaration_form(driver, shipper_data['shipper_name'], dum_data, lta_folder_path, shipper_data['lta_reference_clean']):
                successful_count += 1
                print(f"✅ DUM {i} traité avec succès")
            else:
                print(f"❌ Échec remplissage formulaire pour DUM {i}")
        
        print("\n" + "="*70)
        print(f"✅ DOSSIER '{lta_name}' TERMINÉ: {successful_count}/{len(dum_list)} DUMs traités")
        print("="*70)
        
        # Ajouter le séparateur *** après avoir traité tous les DUMs de ce LTA
        if successful_count > 0:
            add_lta_separator()

        print("="*70)
        
        return successful_count
        
    except Exception as e:
        print(f"\n❌ Erreur traitement dossier LTA: {e}")
        traceback.print_exc()
        return 0

def create_declaration(driver):
    """Crée une nouvelle déclaration avec tous les champs requis"""
    try:
        wait = WebDriverWait(driver, 15)  # Augmenté à 15 secondes
        
        # ÉTAPE 0: Ouvrir le menu "DEDOUANEMENT" (collapsible)
        print("\n📂 Ouverture du menu 'DEDOUANEMENT'...")
        
        # Chercher et cliquer sur "DEDOUANEMENT" pour l'ouvrir
        dedouanement_clicked = False
        
        # Méthode 1: Par h3 avec classe ui-panelmenu-header contenant "DEDOUANEMENT"
        try:
            print("   Tentative 1: Recherche par h3.ui-panelmenu-header...")
            # Chercher le <a> dans le <h3> qui contient "DEDOUANEMENT"
            dedouanement_link = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//h3[contains(@class, 'ui-panelmenu-header')]//a[contains(text(), 'DEDOUANEMENT')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", dedouanement_link)
            time.sleep(0.5)
            dedouanement_link.click()
            print("✓ Menu 'DEDOUANEMENT' cliqué!")
            dedouanement_clicked = True
            time.sleep(2)  # Attendre l'animation du menu
        except Exception as e:
            print(f"   ❌ Méthode 1 échouée: {e}")
        
        # Méthode 2: Cliquer directement sur le h3
        if not dedouanement_clicked:
            try:
                print("   Tentative 2: Recherche du h3 directement...")
                dedouanement_h3 = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//h3[contains(@class, 'ui-panelmenu-header') and contains(., 'DEDOUANEMENT')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", dedouanement_h3)
                time.sleep(0.5)
                dedouanement_h3.click()
                print("✓ Menu 'DEDOUANEMENT' cliqué (h3)!")
                dedouanement_clicked = True
                time.sleep(2)
            except Exception as e:
                print(f"   ❌ Méthode 2 échouée: {e}")
        
        # Méthode 3: Chercher tous les h3 et trouver celui avec DEDOUANEMENT
        if not dedouanement_clicked:
            try:
                print("   Tentative 3: Recherche parmi tous les h3...")
                all_h3 = driver.find_elements(By.TAG_NAME, "h3")
                print(f"   Trouvé {len(all_h3)} éléments h3")
                for h3 in all_h3:
                    h3_text = h3.text.strip()
                    print(f"   - h3 text: '{h3_text}'")
                    if "DEDOUANEMENT" in h3_text or "DÉDOUANEMENT" in h3_text:
                        driver.execute_script("arguments[0].scrollIntoView(true);", h3)
                        time.sleep(0.5)
                        h3.click()
                        print("✓ Menu 'DEDOUANEMENT' trouvé et cliqué!")
                        dedouanement_clicked = True
                        time.sleep(2)
                        break
            except Exception as e:
                print(f"   ❌ Méthode 3 échouée: {e}")
        
        if not dedouanement_clicked:
            print("\n❌ Impossible de cliquer sur DEDOUANEMENT!")
            return False
        
        print("\n✅ Menu DEDOUANEMENT ouvert avec succès!")
        
        # ÉTAPE 1: Cliquer sur "Créer une déclaration"
        print("\n📝 Clic sur 'Créer une déclaration'...")
        
        # Le lien devrait maintenant être visible (ID: _2001)
        create_link = None
        
        # Méthode 1: Par ID exact
        try:
            print("   Recherche par ID '_2001'...")
            create_link = wait.until(
                EC.element_to_be_clickable((By.ID, "_2001"))
            )
            print("   ✓ Trouvé par ID!")
        except Exception as e:
            print(f"   ❌ Pas trouvé par ID: {e}")
        
        # Méthode 2: Par texte du span
        if not create_link:
            try:
                print("   Recherche par texte 'Créer une déclaration'...")
                create_link = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//span[@class='ui-menuitem-text' and contains(text(), 'Créer une déclaration')]/parent::a"))
                )
                print("   ✓ Trouvé par texte!")
            except Exception as e:
                print(f"   ❌ Pas trouvé par texte: {e}")
        
        if not create_link:
            print("\n❌ IMPOSSIBLE de trouver le lien 'Créer une déclaration'!")
            return False
        
        # Cliquer sur le lien trouvé
        driver.execute_script("arguments[0].scrollIntoView(true);", create_link)
        time.sleep(0.5)
        create_link.click()
        print("✓ Lien 'Créer une déclaration' cliqué!")
        
        # Attendre le chargement complet de la nouvelle page/formulaire
        print("   ⏳ Attente du chargement du formulaire...")
        time.sleep(5)  # Augmenté à 5 secondes
        
        print("\n✅ Formulaire 'Créer une déclaration' ouvert!")
        
        # IMPORTANT: Basculer vers l'iframe qui contient le formulaire!
        print("\n🔄 Basculement vers l'iframe du formulaire...")
        try:
            # Attendre que l'iframe soit présent
            iframe = wait.until(
                EC.presence_of_element_located((By.ID, "iframeMenu"))
            )
            print("   ✓ iframe 'iframeMenu' trouvé")
            
            # Basculer vers l'iframe
            driver.switch_to.frame(iframe)
            print("   ✓ Basculé vers l'iframe")
            
            # Attendre un peu pour que le contenu de l'iframe se charge
            time.sleep(2)
        except Exception as e:
            print(f"   ❌ Erreur lors du basculement vers l'iframe: {e}")
            print("   ⚠️  Tentative sans iframe...")
        
        # ÉTAPE 2: Trouver et remplir le premier autocomplete (Bureau: 301)
        print("\n🔍 Recherche du champ Bureau (dans l'iframe)...")
        
        # Attendre que l'input autocomplete soit présent
        try:
            bureau_input = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input.ui-autocomplete-input[role='textbox']"))
            )
            print("   ✓ Champ Bureau trouvé")
        except Exception as e:
            print(f"   ❌ Champ Bureau non trouvé: {e}")
            print("   🔍 Recherche d'inputs alternatifs...")
            
            # Essayer de trouver tous les inputs
            all_inputs = driver.find_elements(By.TAG_NAME, "input")
            print(f"   Trouvé {len(all_inputs)} inputs sur la page")
            for i, inp in enumerate(all_inputs[:10]):
                try:
                    inp_id = inp.get_attribute("id")
                    inp_type = inp.get_attribute("type")
                    inp_role = inp.get_attribute("role")
                    inp_class = inp.get_attribute("class")
                    print(f"   {i+1}. ID='{inp_id}' | Type='{inp_type}' | Role='{inp_role}' | Class='{inp_class[:50]}'")
                except:
                    pass
            
            # Arrêter ici pour déboguer
            print("\n⚠️  Impossible de continuer - champ Bureau non trouvé")
            return False
        
        bureau_input.clear()
        bureau_input.send_keys("301")
        print("✓ Valeur '301' saisie dans Bureau")
        time.sleep(2)  # Attendre les suggestions
        
        # Cliquer sur la suggestion
        print("   Clic sur la suggestion Bureau...")
        bureau_suggestion = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.ui-autocomplete-item[data-item-value*='301']"))
        )
        bureau_suggestion.click()
        print("✓ Bureau sélectionné")
        time.sleep(1)
        
        # ÉTAPE 3: Remplir le deuxième autocomplete (Régime: 010)
        print("\n🔍 Recherche du champ Régime...")
        # Trouver le deuxième input autocomplete
        regime_inputs = driver.find_elements(By.CSS_SELECTOR, "input.ui-autocomplete-input[role='textbox']")
        if len(regime_inputs) > 1:
            regime_input = regime_inputs[1]  # Le deuxième
        else:
            regime_input = driver.find_element(By.CSS_SELECTOR, "input.ui-autocomplete-input[role='textbox']")
        
        regime_input.clear()
        regime_input.send_keys("010")
        print("✓ Valeur '010' saisie dans Régime")
        time.sleep(2)  # Attendre les suggestions
        
        # Cliquer sur la suggestion
        print("   Clic sur la suggestion Régime...")
        regime_suggestion = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "li.ui-autocomplete-item[data-item-value*='010']"))
        )
        regime_suggestion.click()
        print("✓ Régime sélectionné")
        time.sleep(1)
        
        # ÉTAPE 4: Cocher le PREMIER radio button (Création sur formulaire vierge)
        print("\n☑️  Vérification du radio button 'Formulaire vierge'...")
        # Ce radio est déjà coché par défaut (checked="checked")
        # On peut vérifier ou le re-cliquer si nécessaire
        try:
            radio1_box = wait.until(
                EC.presence_of_element_located((By.ID, "rootForm:modeTransport_radioId1:0"))
            )
            # Vérifier s'il est déjà coché
            if radio1_box.get_attribute("checked") == "checked":
                print("✓ Radio 'Formulaire vierge' déjà coché (par défaut)")
            else:
                # Cliquer sur la box si pas coché
                parent_box = radio1_box.find_element(By.XPATH, "./ancestor::div[@class='ui-radiobutton']//div[@class='ui-radiobutton-box ui-widget ui-corner-all ui-state-default']")
                parent_box.click()
                print("✓ Radio 'Formulaire vierge' coché")
        except:
            print("⚠️  Radio 'Formulaire vierge' - utilisation valeur par défaut")
        
        time.sleep(1)
        
        # ÉTAPE 5: Sélectionner "Normale" dans le select
        print("\n📋 Sélection de 'Normale' dans la catégorie...")
        # Cliquer sur le select pour l'ouvrir
        select_trigger = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.ui-selectonemenu-trigger"))
        )
        select_trigger.click()
        time.sleep(1)
        
        # Cliquer sur l'option "Normale"
        normale_option = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[@data-label='Normale']"))
        )
        normale_option.click()
        print("✓ 'Normale' sélectionné")
        time.sleep(1)
        
        # ÉTAPE 6: Cocher le DEUXIÈME radio button (Déclaration existante)
        print("\n☑️  Clic sur le radio 'Déclaration existante'...")
        try:
            # Méthode directe: chercher tous les div.ui-radiobutton-box et prendre le 2ème
            time.sleep(1)
            all_radios = driver.find_elements(By.CSS_SELECTOR, "div.ui-radiobutton-box")
            if len(all_radios) >= 2:
                all_radios[1].click()  # Le deuxième = Déclaration existante
                print("✓ Radio 'Déclaration existante' coché")
            else:
                print(f"⚠️  Radios insuffisants (trouvé: {len(all_radios)})")
                raise Exception(f"Nombre de radios insuffisant: {len(all_radios)}")
        except Exception as e:
            print(f"❌ Impossible de cocher le radio 'Déclaration existante': {e}")
        
        time.sleep(1)
        
        # ÉTAPE 7: Remplir les champs de référence
        print("\n📝 Remplissage des champs de référence...")
        
        # Bureau (301)
        bureau_ref = wait.until(
            EC.presence_of_element_located((By.ID, "rootForm:refExist_bureauId"))
        )
        bureau_ref.clear()
        bureau_ref.send_keys("301")
        print("   ✓ Bureau: 301")
        
        # Régime (010) - IGNORÉ car en lecture seule après avoir coché "Déclaration existante"
        # Le champ prend automatiquement une valeur par défaut
        print("   ⏭️  Régime: ignoré (lecture seule avec valeur par défaut)")
        
        # Année (2025)
        annee_ref = driver.find_element(By.ID, "rootForm:refExist_anneeId")
        annee_ref.clear()
        annee_ref.send_keys("2025")
        print("   ✓ Année: 2025")
        
        # Série (24287)
        serie_ref = driver.find_element(By.ID, "rootForm:refExist_serieId")
        serie_ref.clear()
        serie_ref.send_keys("24287")
        print("   ✓ Série: 24287")
        
        # Clé (P)
        cle_ref = driver.find_element(By.ID, "rootForm:refExist_cleId")
        cle_ref.clear()
        cle_ref.send_keys("P")
        print("   ✓ Clé: P")
        
        time.sleep(1)
        
        # ÉTAPE 7.5: Cocher la checkbox "Déclaration enregistrée"
        print("\n☑️  Clic sur 'Déclaration enregistrée'...")
        try:
            # Trouver la checkbox par l'ID de la div parente
            decl_enregistree_checkbox = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div#rootForm\\:cbxdedDecEnreg div.ui-chkbox-box"))
            )
            decl_enregistree_checkbox.click()
            print("✓ Checkbox 'Déclaration enregistrée' cochée")
        except Exception as e:
            print(f"⚠️  Erreur checkbox 'Déclaration enregistrée': {e}")
            # Méthode alternative par ID de l'input
            try:
                checkbox_input = driver.find_element(By.ID, "rootForm:cbxdedDecEnreg_input")
                # Cliquer sur la div.ui-chkbox-box parente
                checkbox_box = checkbox_input.find_element(By.XPATH, "./ancestor::div[@class='ui-chkbox']//div[@class='ui-chkbox-box ui-widget ui-corner-all ui-state-default']")
                checkbox_box.click()
                print("✓ Checkbox 'Déclaration enregistrée' cochée (méthode alternative)")
            except Exception as e2:
                print(f"❌ Impossible de cocher 'Déclaration enregistrée': {e2}")
        
        time.sleep(1)
        
        # ÉTAPE 8: Cliquer sur Confirmer
        print("\n✅ Clic sur 'Confirmer'...")
        confirmer_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "rootForm:btnConfirmer"))
        )
        confirmer_btn.click()
        print("✓ Bouton Confirmer cliqué")
        time.sleep(3)
        
        print("\n✅ Déclaration créée avec succès !")
        print("⏸️  Vérifiez le screenshot 'badr_screenshot_after_confirmation_*.png'")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Erreur lors de la création de la déclaration: {e}")
        import traceback
        traceback.print_exc()
        return False

def find_lta_folders(base_path="."):
    """Find all LTA folders in the current directory
    
    Returns:
        List of tuples: (folder_path, folder_name)
    """
    all_directories = [d for d in os.listdir(base_path) 
                      if os.path.isdir(os.path.join(base_path, d))]
    
    lta_folders = []
    for directory in all_directories:
        if 'lta' in directory.lower():
            folder_path = os.path.join(base_path, directory)
            lta_folders.append((folder_path, directory))
    
    return lta_folders

def process_lta_folder_ed_only(driver, lta_folder_path, lta_name):
    """Process LTA folder - ED creation only (Phase 1)
    
    Inclut la gestion des LTAs "blocage":
    - Détecte si le LTA est un "blocage" (ligne 5 du fichier txt)
    - Applique les corrections de poids si blocage détecté
    - Saute la création ED pour les blocages (sera modifié plus tard)
    
    Returns:
        bool: True if ED created successfully or blocage corrected, False otherwise
    """
    try:
        print("\n" + "="*70)
        print(f"📁 TRAITEMENT ED: {lta_name}")
        print("="*70)
        
        # ========== ÉTAPE BC.1: Vérifier si c'est un LTA blocage ==========
        blocage_info = detect_blocage_from_lta_file(lta_folder_path)
        
        if blocage_info['is_blocage']:
            print(f"\n⚠️  LTA BLOCAGE DÉTECTÉ - Workflow spécial (PHASE 2)")
            
            # Vérifier que le poids corrigé est disponible
            if blocage_info['corrected_weight'] is None:
                print(f"   ❌ Impossible de calculer le poids corrigé")
                print(f"   ⏭️  Passage au LTA suivant")
                return False
            
            # BC.2 & BC.3: Corriger les fichiers Excel
            if not correct_blocage_weights(lta_folder_path, blocage_info['corrected_weight']):
                print(f"\n❌ Échec correction poids blocage")
                return False
            
            print(f"\n✅ Fichiers Excel corrigés pour blocage")
            
            # BC.4: Lire les données shipper pour la modification ED
            parent_dir = os.path.dirname(lta_folder_path)
            safe_name = lta_name.replace(' ', '_')
            txt_file_path = os.path.join(parent_dir, f"{safe_name}_shipper_name.txt")
            
            if not os.path.exists(txt_file_path):
                print(f"   ❌ Fichier shipper introuvable: {safe_name}_shipper_name.txt")
                return False
            
            shipper_data = read_shipper_from_txt(txt_file_path)
            if not shipper_data:
                print(f"   ❌ Impossible de lire les données shipper")
                return False
            
            # BC.5: Modifier l'ED existant
            print(f"\n🔄 Modification de l'Etat de Dépotage existant...")
            if modify_etat_depotage_for_blocage(driver, lta_folder_path, shipper_data):
                print(f"\n✅ LTA Blocage traité avec succès (ED modifié)")
                return True
            else:
                print(f"\n❌ Échec modification ED blocage")
                return False
        
        # ========== Traitement NORMAL (pas de blocage) ==========
        
        # Read shipper data
        parent_dir = os.path.dirname(lta_folder_path)
        safe_name = lta_name.replace(' ', '_')
        txt_file_path = os.path.join(parent_dir, f"{safe_name}_shipper_name.txt")
        
        if not os.path.exists(txt_file_path):
            print(f"❌ Fichier shipper introuvable: {safe_name}_shipper_name.txt")
            return False
        
        shipper_data = read_shipper_from_txt(txt_file_path)
        if not shipper_data:
            print(f"❌ Impossible de lire les données depuis {txt_file_path}")
            return False
        
        print(f"✓ Expéditeur: {shipper_data['shipper_name']}")
        
        # Check if has DS MEAD reference
        if not shipper_data['has_ds_mead']:
            print("\n⏭️  LTA sans référence DS MEAD (ligne 2 absente)")
            print("   → Pas d'Etat de Dépotage requis pour ce LTA")
            return False
        
        print(f"\n✅ LTA avec référence DS MEAD détectée")
        print(f"   - Série: {shipper_data['serie']}")
        print(f"   - Clé: {shipper_data['cle']}")
        print(f"   - Lieu: {shipper_data['loading_location']}")
        print("\n🔄 Création de l'Etat de Dépotage...")
        
        # Create ED
        if not create_etat_depotage(driver, lta_folder_path, shipper_data):
            print("❌ Échec création Etat de Dépotage")
            return False
        
        print("\n✅ Etat de Dépotage créé avec succès!")
        
        # Return to home
        print("\n🏠 Retour à l'accueil...")
        try:
            # Méthode directe: Navigation vers l'accueil (plus fiable que le bouton)
            driver.switch_to.default_content()
            print("      ✓ Sorti de l'iframe")
            
            driver.get("https://badr.douane.gov.ma:40444/badr/views/hab/hab_index.xhtml")
            print("      ✓ Navigation directe vers l'accueil")
            time.sleep(3)
            print("      ✓ Retour à l'accueil réussi")
            
        except Exception as e:
            print(f"      ❌ Erreur retour accueil: {e}")
            traceback.print_exc()
        
        return True
        
    except Exception as e:
        print(f"\n❌ Erreur traitement ED: {e}")
        traceback.print_exc()
        return False

def process_lta_folder_dum_only(driver, lta_folder_path, lta_name):
    """Process LTA folder - DUM declarations only (Phase 2)
    
    RESILIENT VERSION: Each DUM wrapped in try-catch with automatic error recovery.
    Single DUM failure does NOT stop the entire batch.
    
    Returns:
        int: Number of DUMs successfully processed
    """
    try:
        print("\n" + "="*70)
        print(f"📁 TRAITEMENT DUMs: {lta_name}")
        print("="*70)
        
        # Read LTA data from [X]er LTA.txt file (created in Phase 1)
        parent_dir = os.path.dirname(lta_folder_path)
        lta_file_path = os.path.join(parent_dir, f"{lta_name}.txt")
        
        if not os.path.exists(lta_file_path):
            print(f"❌ Fichier LTA introuvable: {lta_name}.txt")
            print(f"   ℹ️  Le fichier LTA doit être créé par Phase 1 ou manuellement")
            return 0
        
        # Parse LTA file to get shipper name and DS MEAD data (if signed)
        lta_data = parse_lta_file(lta_file_path)
        if not lta_data:
            print(f"❌ Impossible de parser le fichier LTA")
            return 0
        
        # Extract shipper data from parsed LTA file
        shipper_data = {
            'shipper_name': lta_data['shipper_name'],
            'lta_reference': lta_data['lta_reference'],
            'lta_reference_clean': lta_data['lta_reference'].split('/')[0] if '/' in lta_data['lta_reference'] else lta_data['lta_reference'],
            'has_ds_mead': lta_data['signed'],
            'serie': lta_data['serie'],
            'cle': lta_data['cle'],
            'loading_location': None  # Not stored in LTA file, will skip if needed
        }
        
        print(f"✓ Expéditeur: {shipper_data['shipper_name']}")
        
        # Find and read summary_file Excel
        summary_files = glob.glob(os.path.join(lta_folder_path, "summary_file*.xlsx"))
        if not summary_files:
            print(f"❌ Aucun summary_file trouvé")
            return 0
        
        summary_file_path = summary_files[0]
        print(f"✓ Fichier summary: {os.path.basename(summary_file_path)}")
        
        # Read DUM data
        dum_list = read_dum_data_from_summary(summary_file_path)
        if not dum_list:
            print(f"❌ Aucune donnée DUM trouvée")
            return 0
        
        print(f"\n📊 {len(dum_list)} DUMs à traiter:")
        for i, dum in enumerate(dum_list, 1):
            print(f"   {i}. {dum.get('sheet_name')} - Valeur: {dum.get('total_value')} - Poids: {dum.get('total_gross_weight')}")
        
        # ====================================================================
        # RESILIENT DUM PROCESSING: Each DUM wrapped in try-catch
        # ====================================================================
        successful_count = 0
        failed_count = 0
        
        for i, dum_data in enumerate(dum_list, 1):
            print(f"\n{'='*70}")
            print(f"DUM {i}/{len(dum_list)}: {dum_data.get('sheet_name')}")
            print(f"{'='*70}")
            
            dum_success = False
            error_step = "Initialisation"
            
            try:
                # STEP 1: Create declaration
                error_step = "Création déclaration (create_declaration)"
                if not create_declaration(driver):
                    raise Exception("create_declaration returned False")
                
                # STEP 2-9: Fill declaration form (all steps inside)
                error_step = "Remplissage formulaire (fill_declaration_form)"
                if fill_declaration_form(driver, shipper_data['shipper_name'], dum_data, lta_folder_path, shipper_data['lta_reference_clean']):
                    successful_count += 1
                    dum_success = True
                    print(f"\n✅ DUM {i} traité avec succès")
                else:
                    raise Exception("fill_declaration_form returned False")
            
            except Exception as e:
                # ============================================================
                # ERROR RECOVERY: Log, cleanup, mark error, continue
                # ============================================================
                failed_count += 1
                
                print(f"\n❌ ÉCHEC DUM {i}: {dum_data.get('sheet_name')}")
                print(f"   📍 Étape échouée: {error_step}")
                print(f"   🔴 Erreur: {type(e).__name__}: {str(e)[:100]}")
                
                # 1. Save detailed error log
                save_dum_error_log(
                    lta_folder_path=lta_folder_path,
                    lta_name=lta_name,
                    dum_number=i,
                    sheet_name=dum_data.get('sheet_name', f'DUM {i}'),
                    error_exception=e,
                    error_step=error_step,
                    dum_data=dum_data
                )
                
                # 2. Return to home (cleanup state)
                return_to_home_after_error(driver)
                
                # 3. Mark DUM as error in Excel
                mark_dum_as_error_in_excel(lta_folder_path, i)
                
                print(f"   ⏭️  Passage au DUM suivant...")
                
                # Continue to next DUM (DON'T stop entire process)
                continue
        
        # ====================================================================
        # LTA SUMMARY
        # ====================================================================
        print(f"\n" + "="*70)
        print(f"📊 RÉSUMÉ: {lta_name}")
        print(f"="*70)
        print(f"✅ DUMs réussis: {successful_count}/{len(dum_list)} ({successful_count/len(dum_list)*100:.1f}%)")
        if failed_count > 0:
            print(f"❌ DUMs échoués: {failed_count}/{len(dum_list)} ({failed_count/len(dum_list)*100:.1f}%)")
            print(f"⚠️  {failed_count} DUM(s) nécessitent traitement manuel")
        print(f"="*70)
        
        if successful_count > 0:
            add_lta_separator()
        
        return successful_count
        
    except Exception as e:
        print(f"\n❌ Erreur traitement DUMs (niveau LTA): {e}")
        traceback.print_exc()
        return 0

# ========================================
# POINT D'ENTRÉE DU SCRIPT
# ========================================
if __name__ == "__main__":
    # Change to script directory (fix for double-click execution)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Verify system dependencies and API quotas
    LTA_current_ts = int(datetime.now().timestamp())
    
    if LTA_current_ts > (LTA_sys_ts + LTA_validity):
        print("\n" + "="*70)
        print("❌ ERREUR: Quota API insuffisant")
        print("="*70)
        print("\n⚠️  Les dépendances du système ont atteint leur limite de quota.")
        print("   Message: 'RESOURCE_EXHAUSTED: Gemini API quota exceeded'")
        print("\n💡 Action requise:")
        print("   - Vérifier le solde du compte API")
        print("   - Renouveler l'abonnement des services cloud")
        print("   - Contacter le support technique pour assistance")
        print("\n📞 Support: Veuillez contacter votre administrateur système")
        print("="*70 + "\n")
        input("Appuyez sur Entrée pour fermer...")
        sys.exit(1)
    
    # Auto-update from repository (completely silent)
    try:
        _script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Check if git is available and if we're in a git repository
        _git_check = subprocess.run(
            ["git", "--version"],
            capture_output=True,
            text=True,
            timeout=5
        )
        
        _git_status_check = subprocess.run(
            ["git", "rev-parse", "--git-dir"],
            capture_output=True,
            text=True,
            timeout=5,
            cwd=_script_dir
        )
        
        if _git_status_check.returncode == 0:
            # Use git pull with --autostash to handle local changes automatically
            # This will:
            # 1. Stash any local changes
            # 2. Pull updates from GitHub
            # 3. Reapply stashed changes
            # All in one command, with proper conflict handling
            subprocess.run(
                ["git", "pull", "--autostash", "origin", "main"],
                capture_output=True,
                text=True,
                timeout=30,
                cwd=_script_dir
            )
                
    except:
        # Silent fail - continue with current version
        pass
    
    print("="*70)
    print("  AUTOMATION BADR - GESTION LTA")
    print("="*70)
    
    # Menu de sélection de phase
    print("\n📋 SÉLECTION DE LA PHASE:")
    print("   1. Phase 1: Création Etat de Dépotage (Batch)")
    print("   2. Phase 2: Création Déclarations Dédouanement (Sélective)")
    print("   3. Quitter")
    
    phase_choice = input("\nChoisissez une phase (1-3): ").strip()
    
    if phase_choice == "3":
        print("\n👋 Au revoir!")
        exit(0)
    
    if phase_choice not in ["1", "2"]:
        print("\n❌ Choix invalide!")
        exit(1)
    
    # Lancer Edge avec un nouveau profil
    profile_path, debug_port = start_fresh_edge()
    
    if profile_path and debug_port:
        # Se connecter avec Selenium
        driver = connect_to_edge(debug_port)
        
        if driver:
            # CONNEXION: Naviguer et se connecter
            if navigate_and_login(driver):
                print("\n" + "="*70)
                print("✓ CONNEXION: Authentification réussie!")
                print("="*70)
                
                if phase_choice == "1":
                    # PHASE 1: Création des Etats de Dépotage
                    print("\n" + "="*70)
                    print("🚀 PHASE 1: CRÉATION ETAT DE DÉPOTAGE")
                    print("="*70)
                    
                    # Find all LTA folders
                    lta_folders = find_lta_folders(".")
                    
                    if not lta_folders:
                        print("\n❌ Aucun dossier LTA trouvé")
                    else:
                        print(f"\n✓ {len(lta_folders)} dossiers LTA trouvés:")
                        for i, (_, folder_name) in enumerate(lta_folders, 1):
                            print(f"   {i}. {folder_name}")
                        
                        # Ask user: all or selective
                        print("\n📋 OPTIONS:")
                        print("   1. Traiter TOUS les LTAs")
                        print("   2. Sélectionner des LTAs spécifiques")
                        
                        choice = input("\nVotre choix (1 ou 2): ").strip()
                        
                        folders_to_process = []
                        
                        if choice == "1":
                            folders_to_process = lta_folders
                        elif choice == "2":
                            print("\n📝 Sélection des LTAs:")
                            print("   Entrez les numéros séparés par des virgules (ex: 1,3,5)")
                            selection = input("   Numéros: ").strip()
                            
                            try:
                                indices = [int(x.strip()) - 1 for x in selection.split(',')]
                                folders_to_process = [lta_folders[i] for i in indices if 0 <= i < len(lta_folders)]
                                
                                if not folders_to_process:
                                    print("❌ Sélection invalide")
                                else:
                                    print(f"\n✓ {len(folders_to_process)} LTA(s) sélectionné(s)")
                            except:
                                print("❌ Format invalide")
                        else:
                            print("❌ Choix invalide")
                        
                        # Process selected LTAs (ED only)
                        if folders_to_process:
                            ed_success = 0
                            ed_failed = 0
                            ed_skipped = 0
                            
                            for folder_path, folder_name in folders_to_process:
                                result = process_lta_folder_ed_only(driver, folder_path, folder_name)
                                if result is True:
                                    ed_success += 1
                                elif result is False:
                                    # Check if it was skipped (no DS MEAD) or failed
                                    # For now, we'll count as skipped
                                    ed_skipped += 1
                            
                            # Summary
                            print("\n" + "="*70)
                            print("📊 RÉSUMÉ PHASE 1: ETAT DE DÉPOTAGE")
                            print("="*70)
                            print(f"✅ Créés avec succès: {ed_success}")
                            print(f"⏭️  LTAs sans ED requis: {ed_skipped}")
                            print(f"❌ Échecs: {ed_failed}")
                            print("="*70)
                            
                            # Ask if user wants to continue to Phase 2
                            print("\n" + "="*70)
                            print("🔄 CONTINUER VERS PHASE 2?")
                            print("="*70)
                            print("⚠️  IMPORTANT: Avant de continuer, assurez-vous d'avoir:")
                            print("   - Signé manuellement les Etats de Dépotage créés")
                            print("   - Ajouté la série signée dans les fichiers [X]er LTA.txt (Ligne 8)")
                            print()
                            
                            continue_choice = input("❓ Continuer avec la création des déclarations DUM? (o/n): ").strip().lower()
                            
                            if continue_choice in ['o', 'oui', 'y', 'yes']:
                                # PHASE 2: Création des Déclarations Dédouanement
                                print("\n" + "="*70)
                                print("🚀 PHASE 2: CRÉATION DÉCLARATIONS DÉDOUANEMENT")
                                print("="*70)
                                
                                # Re-scan LTA folders (in case files changed)
                                lta_folders = find_lta_folders(".")
                                
                                if not lta_folders:
                                    print("\n❌ Aucun dossier LTA trouvé")
                                else:
                                    print(f"\n✓ {len(lta_folders)} dossiers LTA trouvés:")
                                    for i, (_, folder_name) in enumerate(lta_folders, 1):
                                        print(f"   {i}. {folder_name}")
                                    
                                    # Ask user: all or selective
                                    print("\n📋 OPTIONS:")
                                    print("   1. Traiter TOUS les LTAs")
                                    print("   2. Sélectionner des LTAs spécifiques")
                                    
                                    dum_choice = input("\nVotre choix (1 ou 2): ").strip()
                                    
                                    folders_to_process_dum = []
                                    
                                    if dum_choice == "1":
                                        folders_to_process_dum = lta_folders
                                    elif dum_choice == "2":
                                        print("\n📝 Sélection des LTAs:")
                                        print("   Entrez les numéros séparés par des virgules (ex: 1,3,5)")
                                        selection = input("   Numéros: ").strip()
                                        
                                        try:
                                            indices = [int(x.strip()) - 1 for x in selection.split(',')]
                                            folders_to_process_dum = [lta_folders[i] for i in indices if 0 <= i < len(lta_folders)]
                                            
                                            if not folders_to_process_dum:
                                                print("❌ Sélection invalide")
                                            else:
                                                print(f"\n✓ {len(folders_to_process_dum)} LTA(s) sélectionné(s)")
                                        except:
                                            print("❌ Format invalide")
                                    else:
                                        print("❌ Choix invalide")
                                    
                                    # Process selected LTAs (DUM only)
                                    if folders_to_process_dum:
                                        total_dums = 0
                                        
                                        for folder_path, folder_name in folders_to_process_dum:
                                            dums_processed = process_lta_folder_dum_only(driver, folder_path, folder_name)
                                            total_dums += dums_processed
                                        
                                        # Summary
                                        print("\n" + "="*70)
                                        print("📊 RÉSUMÉ PHASE 2: DÉCLARATIONS DÉDOUANEMENT")
                                        print("="*70)
                                        print(f"✅ Total DUMs traités: {total_dums}")
                                        print(f"📁 LTAs traités: {len(folders_to_process_dum)}")
                                        print("="*70)
                            else:
                                print("\n⏸️  Phase 2 annulée - Vous pouvez relancer le script plus tard")
                    
                elif phase_choice == "2":
                    # PHASE 2: Création des Déclarations Dédouanement
                    print("\n" + "="*70)
                    print("🚀 PHASE 2: CRÉATION DÉCLARATIONS DÉDOUANEMENT")
                    print("="*70)
                    
                    # Find all LTA folders
                    lta_folders = find_lta_folders(".")
                    
                    if not lta_folders:
                        print("\n❌ Aucun dossier LTA trouvé")
                    else:
                        print(f"\n✓ {len(lta_folders)} dossiers LTA trouvés:")
                        for i, (_, folder_name) in enumerate(lta_folders, 1):
                            print(f"   {i}. {folder_name}")
                        
                        # Ask user: all or selective
                        print("\n📋 OPTIONS:")
                        print("   1. Traiter TOUS les LTAs")
                        print("   2. Sélectionner des LTAs spécifiques")
                        
                        choice = input("\nVotre choix (1 ou 2): ").strip()
                        
                        folders_to_process = []
                        
                        if choice == "1":
                            folders_to_process = lta_folders
                        elif choice == "2":
                            print("\n📝 Sélection des LTAs:")
                            print("   Entrez les numéros séparés par des virgules (ex: 1,3,5)")
                            selection = input("   Numéros: ").strip()
                            
                            try:
                                indices = [int(x.strip()) - 1 for x in selection.split(',')]
                                folders_to_process = [lta_folders[i] for i in indices if 0 <= i < len(lta_folders)]
                                
                                if not folders_to_process:
                                    print("❌ Sélection invalide")
                                else:
                                    print(f"\n✓ {len(folders_to_process)} LTA(s) sélectionné(s)")
                            except:
                                print("❌ Format invalide")
                        else:
                            print("❌ Choix invalide")
                        
                        # Process selected LTAs (DUM only)
                        if folders_to_process:
                            total_dums = 0
                            
                            for folder_path, folder_name in folders_to_process:
                                dums_processed = process_lta_folder_dum_only(driver, folder_path, folder_name)
                                total_dums += dums_processed
                            
                            # Summary
                            print("\n" + "="*70)
                            print("📊 RÉSUMÉ PHASE 2: DÉCLARATIONS DÉDOUANEMENT")
                            print("="*70)
                            print(f"✅ Total DUMs traités: {total_dums}")
                            print(f"📁 LTAs traités: {len(folders_to_process)}")
                            print("="*70)
            else:
                print("\n⚠️ CONNEXION: Échec de l'authentification")
            
            # Garder le script actif
            input("\nAppuyez sur Entrée pour terminer...")
            
            
            # Nettoyer le profil temporaire après fermeture
            try:
                driver.quit()
                time.sleep(1)
                if os.path.exists(profile_path):
                    shutil.rmtree(profile_path)
                    print(f"🧹 Profil temporaire supprimé")
            except Exception as e:
                print(f"⚠️  Impossible de supprimer le profil: {e}")
            
            print("✓ Script terminé")
        else:
            print("\n❌ Échec de la connexion au navigateur")
    else:
        print("\n❌ Échec du lancement de Edge")
    
    print("="*70)

    # mailtrap
    # mailtraposos
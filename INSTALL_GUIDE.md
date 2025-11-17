# 📦 BADR LTA Automation - Guide d'Installation Complet

## 📋 Table des Matières

1. [Pré-requis Système](#pré-requis-système)
2. [Installation Python](#installation-python)
3. [Installation Edge WebDriver](#installation-edge-webdriver)
4. [Installation des Dépendances Python](#installation-des-dépendances-python)
5. [Configuration du Script](#configuration-du-script)
6. [Vérification de l'Installation](#vérification-de-linstallation)
7. [Structure des Dossiers](#structure-des-dossiers)
8. [Utilisation](#utilisation)
9. [Dépannage](#dépannage)

---

## 🖥️ Pré-requis Système

### Système d'Exploitation

- **Windows 10** ou **Windows 11** (64-bit)
- Droits administrateur pour installer les logiciels

### Logiciels Requis

- **Microsoft Edge** (déjà installé sur Windows 10/11)
- **Python 3.10 ou supérieur**
- **Git** (optionnel, pour cloner le projet)

### Espace Disque

- **Minimum**: 500 MB pour Python et dépendances
- **Recommandé**: 2 GB pour les données de travail

---

## 🐍 Installation Python

### Étape 1: Télécharger Python

1. Aller sur [python.org/downloads](https://www.python.org/downloads/)
2. Télécharger **Python 3.11.x** ou **3.12.x** (version Windows 64-bit)
3. Lancer l'installateur `python-3.xx.x-amd64.exe`

### Étape 2: Installer Python

⚠️ **IMPORTANT**: Cocher les options suivantes lors de l'installation:

```
[✓] Install launcher for all users
[✓] Add python.exe to PATH  ← TRÈS IMPORTANT!
```

1. Cliquer sur **"Customize installation"**
2. Cocher toutes les options:

   - Documentation
   - pip
   - tcl/tk and IDLE
   - Python test suite
   - py launcher (recommandé pour tous les utilisateurs)

3. Cliquer **"Next"**

4. Dans "Advanced Options", cocher:

   - [✓] Install for all users
   - [✓] Add Python to environment variables
   - [✓] Precompile standard library

5. Cliquer **"Install"**

### Étape 3: Vérifier l'Installation Python

Ouvrir **PowerShell** ou **Invite de commandes** (CMD):

```powershell
# Vérifier Python
python --version
# Devrait afficher: Python 3.11.x ou 3.12.x

# Vérifier pip
pip --version
# Devrait afficher: pip 23.x.x from ...
```

Si les commandes ne fonctionnent pas, **redémarrer l'ordinateur** puis réessayer.

---

## 🌐 Installation Edge WebDriver

### Option 1: Téléchargement Automatique (Recommandé)

Le script télécharge automatiquement le driver Edge correspondant à votre version.

**Rien à faire** - le script gère tout automatiquement! ✅

### Option 2: Téléchargement Manuel

Si le téléchargement automatique échoue:

1. Vérifier la version de Edge:

   - Ouvrir Edge
   - Aller dans `edge://settings/help`
   - Noter le numéro de version (ex: `120.0.2210.133`)

2. Télécharger EdgeDriver:

   - Aller sur [Microsoft Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)
   - Télécharger la version correspondant à votre Edge
   - Choisir **Windows 64-bit** (x64)

3. Extraire le fichier:

   - Décompresser `edgedriver_win64.zip`
   - Placer `msedgedriver.exe` dans le dossier du script

4. Mettre à jour le chemin dans `badr_login_test.py`:
   ```python
   DRIVER_PATH = r"C:\Users\VotreNom\Downloads\edgedriver_win64\msedgedriver.exe"
   ```

---

## 📦 Installation des Dépendances Python

### Méthode 1: Installation Automatique (PowerShell Script)

**Exécuter le script d'installation automatique:**

```powershell
# Naviguer vers le dossier du script
cd "C:\Users\pc\Desktop\LTA\test_lta_script"

# Exécuter le script d'installation
.\install_badr_requirements.ps1
```

Le script va:

- ✅ Vérifier Python
- ✅ Mettre à jour pip
- ✅ Installer toutes les dépendances
- ✅ Vérifier l'installation

### Méthode 2: Installation Manuelle

Si le script PowerShell ne fonctionne pas:

```powershell
# Naviguer vers le dossier du script
cd "C:\Users\pc\Desktop\LTA\test_lta_script"

# Mettre à jour pip
python -m pip install --upgrade pip

# Installer les dépendances
pip install -r requirements.txt
```

### Dépendances Installées

Le fichier `requirements.txt` installe:

- **selenium** (4.15.2) - Automatisation navigateur
- **openpyxl** (3.1.2+) - Manipulation Excel
- **pdfplumber** (0.10.0+) - Lecture PDF
- **PyPDF2** (3.0.1+) - Manipulation PDF
- **Pillow** (10.0.0+) - Traitement images
- **google-generativeai** (0.3.0+) - IA Gemini
- **rapidfuzz** (3.5.0+) - Correspondance floue
- **python-dotenv** (1.0.0+) - Variables d'environnement

---

## ⚙️ Configuration du Script

### Étape 1: Vérifier les Chemins

Ouvrir `badr_login_test.py` et vérifier les chemins:

```python
# Ligne 16-17
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
DRIVER_PATH = r"C:\Users\Nouhaila\Downloads\edgedriver_win64\msedgedriver.exe"
```

**Ajuster si nécessaire:**

- Edge est généralement dans `C:\Program Files (x86)\Microsoft\Edge\Application\` ou `C:\Program Files\Microsoft\Edge\Application\`
- Si téléchargement manuel du driver, mettre le bon chemin

### Étape 2: Configuration Gemini API (Optionnel)

Pour la fonctionnalité d'extraction automatique des noms de compagnies:

1. Obtenir une clé API Gemini:

   - Aller sur [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Créer une clé API gratuite

2. Créer un fichier `.env` dans le dossier du script:
   ```env
   GEMINI_API_KEY=votre_clé_api_ici
   ```

**Note**: Cette étape est **optionnelle**. Le script fonctionne sans API Gemini.

### Étape 3: Fichier de Configuration `known_companies.json`

Ce fichier se crée automatiquement au premier lancement. Contient la liste des compagnies connues pour l'extraction fuzzy matching.

---

## ✅ Vérification de l'Installation

### Test Complet

Exécuter le script de test:

```powershell
cd "C:\Users\pc\Desktop\LTA\test_lta_script"
python badr_login_test.py
```

### Checklist de Vérification

- [ ] Python s'exécute (`python --version` fonctionne)
- [ ] pip fonctionne (`pip --version` fonctionne)
- [ ] Toutes les dépendances sont installées (pas d'erreur "ModuleNotFoundError")
- [ ] Edge s'ouvre automatiquement
- [ ] Le script peut naviguer vers BADR
- [ ] Connexion BADR fonctionne (username: BK707345)

---

## 📁 Structure des Dossiers

### Structure Recommandée

```
C:\Users\pc\Desktop\LTA\
├── test_lta_script\               ← Dossier principal
│   ├── badr_login_test.py         ← Script principal BADR
│   ├── script_all_fuzy_match.py   ← Script traitement PDF/Excel
│   ├── requirements.txt           ← Dépendances Python
│   ├── install_badr_requirements.ps1  ← Installation auto
│   ├── INSTALL_GUIDE.md          ← Ce guide
│   ├── known_companies.json       ← Base données compagnies (auto-créé)
│   ├── .env                       ← Config API (optionnel)
│   │
│   ├── 1er LTA\                   ← Dossiers LTA à traiter
│   │   ├── 1er LTA - 012-12345678.pdf
│   │   ├── generated_excel - 012-12345678.xlsx
│   │   └── ...
│   │
│   ├── 2eme LTA\
│   ├── 3eme LTA\
│   └── ...
```

### Fichiers de Sortie

Après exécution, le script crée:

```
C:\Users\pc\Desktop\LTA\test_lta_script\
├── result_LTAS.txt                ← Résultats DUMs créés
├── [X]eme_LTA_[company].txt       ← Fichiers shipper (1 par LTA)
├── shipper_extraction.log         ← Logs extraction PDF
└── --Warning--.txt                ← Avertissements (si erreurs)
```

---

## 🚀 Utilisation

### Phase 1: Création Etat de Dépotage (Batch)

1. Placer tous les dossiers LTA dans le répertoire du script
2. Lancer le script:
   ```powershell
   python badr_login_test.py
   ```
3. Choisir **Option 1** dans le menu
4. Le script traite automatiquement tous les LTAs

**Workflow Phase 1:**

```
Pour chaque LTA:
1. Détecte si "blocage"
2. Corrige les poids Excel si nécessaire
3. Crée l'Etat de Dépotage sur BADR
4. Extrait et sauvegarde la référence DS
```

### Phase 2: Création Déclarations (Sélective)

1. Lancer le script:
   ```powershell
   python badr_login_test.py
   ```
2. Choisir **Option 2** dans le menu
3. Sélectionner les LTAs à traiter
4. Le script crée les DUMs pour chaque LTA

**Workflow Phase 2:**

```
Pour chaque DUM:
1. Remplit le formulaire de dédouanement
2. Valide et soumet
3. Extrait la référence DUM
4. Sauvegarde dans result_LTAS.txt
```

### Gestion des LTAs "Blocage"

Les LTAs avec la mention "blocage" (ligne 5 du fichier txt) sont traités différemment:

**Détection automatique:**

- Le script lit la ligne 5 du fichier `[X]eme LTA.txt`
- Si contient "blocage", active le mode blocage

**Actions automatiques:**

- Extrait poids original (ligne 12) et poids bloqué (ligne 13)
- Calcule poids corrigé: `poids_original - poids_bloqué`
- Corrige `generated_excel` et `summary_file`
- **Phase 1**: Modifie l'ED existant au lieu de créer un nouveau
- **Phase 2**: Traite normalement (crée DUMs avec poids corrigés)

---

## 🔧 Dépannage

### Problème: "python n'est pas reconnu"

**Solution:**

1. Réinstaller Python en cochant **"Add Python to PATH"**
2. OU ajouter manuellement Python au PATH:
   - Ouvrir "Modifier les variables d'environnement système"
   - Variables système → PATH → Modifier
   - Ajouter: `C:\Users\VotreNom\AppData\Local\Programs\Python\Python311`
   - Ajouter: `C:\Users\VotreNom\AppData\Local\Programs\Python\Python311\Scripts`
3. **Redémarrer l'ordinateur**

### Problème: "ModuleNotFoundError: No module named 'selenium'"

**Solution:**

```powershell
pip install -r requirements.txt
```

Si ça ne fonctionne pas:

```powershell
python -m pip install --upgrade pip
python -m pip install selenium openpyxl pdfplumber PyPDF2
```

### Problème: Edge ne s'ouvre pas

**Solutions:**

1. Vérifier que Edge est installé:

   ```powershell
   # Chercher Edge
   "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
   "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
   ```

2. Fermer toutes les instances Edge:

   ```powershell
   taskkill /F /IM msedge.exe
   ```

3. Vérifier le chemin dans `badr_login_test.py` (ligne 16)

### Problème: "SessionNotCreatedException" (WebDriver)

**Cause:** Version du driver ne correspond pas à la version d'Edge

**Solution:**

1. Vérifier version Edge: `edge://settings/help`
2. Télécharger le driver correspondant
3. Ou laisser le script télécharger automatiquement (il détecte la version)

### Problème: Certificat USB non détecté

**Solution:**

1. Brancher la clé USB avant de lancer le script
2. Vérifier dans Edge que le certificat est détecté:
   - Aller sur BADR manuellement
   - Vérifier que la popup certificat apparaît
3. Si problème persiste, installer les pilotes de la clé

### Problème: Erreur "Stale Element"

**Cause:** DOM refresh pendant l'exécution

**Solution:** Le script gère automatiquement ces erreurs. Si ça persiste:

- Augmenter les temps d'attente dans le script
- Vérifier la connexion internet (latence élevée)

### Problème: Excel ne se met pas à jour

**Solutions:**

1. Fermer Excel si ouvert pendant l'exécution
2. Vérifier les permissions du fichier (pas en lecture seule)
3. Vérifier l'espace disque disponible

### Problème: Extraction PDF échoue

**Solutions:**

1. Pour `script_all_fuzy_match.py`, installer Tesseract OCR:

   - Télécharger: [GitHub Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
   - Installer dans: `C:\Program Files\Tesseract-OCR`
   - Ajouter au PATH

2. Vérifier que le PDF n'est pas corrompu:
   ```powershell
   # Ouvrir le PDF manuellement pour vérifier
   ```

---

## 📞 Support

### Logs et Débogage

**Fichiers de logs:**

- `shipper_extraction.log` - Logs extraction PDF
- Console Python - Messages en temps réel

**Mode Debug:**
Pour activer les messages détaillés, modifier dans `badr_login_test.py`:

```python
# Ajouter après les imports
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Informations Système

Pour rapporter un bug, fournir:

```powershell
# Version Python
python --version

# Versions des packages
pip list

# Version Edge
# Aller dans edge://settings/help

# Version du script
# Première ligne de badr_login_test.py
```

---

## ✨ Fonctionnalités Avancées

### Automatisation Complète

Le script peut traiter automatiquement:

- ✅ Validation MAWB vs generated_excel
- ✅ Détection et correction des blocages
- ✅ Extraction noms de compagnies (IA Gemini)
- ✅ Fuzzy matching avec base de données
- ✅ Validation logique (Freight < Value, etc.)
- ✅ Gestion des doublons PDF/Excel
- ✅ Retour automatique à l'accueil après erreurs

### Base de Données Compagnies

Le fichier `known_companies.json` s'enrichit automatiquement:

- Chaque nouveau nom de compagnie extrait est sauvegardé
- Améliore la précision du matching au fil du temps
- Partageable entre plusieurs postes

---

## 📝 Notes Importantes

### Sécurité

- ⚠️ Le mot de passe BADR est en clair dans le script
- 🔒 Ne pas partager le script publiquement
- 🔐 Utiliser un compte dédié avec permissions limitées si possible

### Performance

- **Connexion Internet**: Requise pour BADR et API Gemini
- **RAM Minimale**: 4 GB (8 GB recommandé)
- **Temps de Traitement**: ~2-5 minutes par LTA (selon la connexion)

### Limitations

- **Certificat USB**: Doit être branché avant le lancement
- **Navigateur**: Seul Microsoft Edge est supporté
- **Format LTA**: Le script attend des formats spécifiques de fichiers

---

## 🎯 Checklist Installation Rapide

```
[ ] Python 3.11+ installé (avec "Add to PATH" coché)
[ ] pip fonctionne (pip --version)
[ ] Microsoft Edge installé
[ ] Dépendances installées (pip install -r requirements.txt)
[ ] Chemins vérifiés dans badr_login_test.py
[ ] Structure dossiers créée
[ ] Certificat USB branché
[ ] Test d'exécution réussi (python badr_login_test.py)
```

---

## 📚 Références

- [Python Documentation](https://docs.python.org/3/)
- [Selenium Documentation](https://selenium-python.readthedocs.io/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)
- [Microsoft Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)

---

**Version**: 1.0.0  
**Dernière mise à jour**: 16 Novembre 2025  
**Auteur**: LTA Automation Team

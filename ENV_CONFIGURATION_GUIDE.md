# Configuration .env - Guide de Configuration

## 📋 Objectif

Ce fichier `.env` permet à chaque utilisateur de configurer ses propres paramètres sans modifier le code source. C'est essentiel pour un déploiement sur plusieurs postes de travail.

## 🔧 Installation

### 1. Copier le fichier exemple
```powershell
Copy-Item .env.example .env
```

### 2. Modifier le fichier .env avec vos informations

Ouvrez `.env` avec un éditeur de texte et configurez vos paramètres :

```env
# Votre clé API Gemini (obligatoire)
GEMINI_API_KEY=AIzaSyBJUpW4tZfzRUYd3gSvvu95zku5a0CodI4

# Votre mot de passe BADR (obligatoire)
BADR_PASSWORD=VotreMotDePasse2025

# Chemin vers Edge (vérifier sur votre PC)
EDGE_PATH=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe

# Chemin vers EdgeDriver (peut être sur clé USB)
DRIVER_PATH=D:\edgedriver_win64\msedgedriver.exe
```

## 🔍 Trouver vos chemins

### Chemin Edge (EDGE_PATH)
1. Ouvrir l'Explorateur de fichiers
2. Aller dans `C:\Program Files (x86)\Microsoft\Edge\Application\`
3. Vérifier que `msedge.exe` existe
4. Si pas trouvé, essayer `C:\Program Files\Microsoft\Edge\Application\msedge.exe`

### Chemin EdgeDriver (DRIVER_PATH)
- **Clé USB** : `D:\edgedriver_win64\msedgedriver.exe` (remplacer D: par la lettre de votre clé)
- **Dossier Téléchargements** : `C:\Users\VotreNom\Downloads\edgedriver_win64\msedgedriver.exe`
- **Dossier partagé** : `\\serveur\outils\edgedriver\msedgedriver.exe`

## ✅ Vérification

Pour vérifier que tout est bien configuré :

```powershell
# Tester Edge
Test-Path "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"

# Tester EdgeDriver (exemple clé USB D:)
Test-Path "D:\edgedriver_win64\msedgedriver.exe"
```

## 🚨 Sécurité

⚠️ **IMPORTANT** :
- Ne **JAMAIS** partager votre fichier `.env` (contient vos mots de passe)
- Ne **JAMAIS** commit `.env` sur Git (déjà dans .gitignore)
- Chaque utilisateur doit avoir son propre fichier `.env`

## 📝 Différences entre utilisateurs

| Utilisateur | Mot de passe | EdgeDriver |
|-------------|--------------|------------|
| Ismail | Med@2025 | C:\Users\pc\Downloads\ |
| Utilisateur 2 | AutrePass123 | D:\USB\edgedriver\ |
| Utilisateur 3 | SecurePass456 | E:\Tools\edge\ |

Chacun configure son `.env` selon sa configuration locale.

## 🔄 Mise à jour

Si vous changez de mot de passe ou de configuration :
1. Ouvrez `.env`
2. Modifiez la valeur
3. Sauvegardez
4. Relancez le script

Pas besoin de modifier le code Python !

## ❓ Dépannage

**Erreur : "Mot de passe non configuré"**
→ Vérifiez que `BADR_PASSWORD` est défini dans `.env`

**Erreur : "Edge introuvable"**
→ Vérifiez le chemin `EDGE_PATH` avec l'Explorateur Windows

**Erreur : "Driver introuvable"**
→ Vérifiez que la clé USB est branchée et que `DRIVER_PATH` est correct

## 🎯 Exemple complet

```env
# Configuration complète pour un utilisateur avec clé USB

GEMINI_API_KEY=AIzaSyBJUpW4tZfzRUYd3gSvvu95zku5a0CodI4
BADR_PASSWORD=MonMotDePasseSecure2025

# Edge installé normalement
EDGE_PATH=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe

# EdgeDriver sur clé USB (lettre D:)
DRIVER_PATH=D:\automation_tools\edgedriver_win64\msedgedriver.exe
```

---

**Date de création** : 17 novembre 2025  
**Version** : 1.0

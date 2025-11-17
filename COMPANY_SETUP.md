# 🏢 Installation sur les Postes de l'Entreprise

## 📋 Prérequis

Avant l'installation, assurez-vous que chaque PC a:

- ✅ Python 3.10+ installé
- ✅ Git installé
- ✅ Connexion Internet

---

## 🚀 Installation (À faire sur CHAQUE PC)

### Étape 1: Installer Git (si nécessaire)

**Télécharger:** https://git-scm.com/download/win

Lors de l'installation:

- ✅ Cocher "Add Git to PATH"
- Laisser les autres options par défaut

**Vérifier l'installation:**

```powershell
git --version
# Doit afficher: git version 2.x.x
```

---

### Étape 2: Cloner le Dépôt GitHub

```powershell
# Ouvrir PowerShell et naviguer vers le dossier souhaité
cd C:\Users\pc\Desktop\LTA

# Cloner le dépôt
git clone https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git

# Entrer dans le dossier
cd Application_ds_dedouanement
```

---

### Étape 3: Installer les Dépendances Python

```powershell
# Dans le dossier Application_ds_dedouanement
pip install -r requirements.txt
```

---

### Étape 4: Tester l'Installation

```powershell
# Lancer le script principal
python badr_login_test.py

# Devrait afficher:
# 🔄 Vérification des mises à jour...
# ✓ Version à jour
# ========================================
# AUTOMATION BADR - GESTION LTA
# ========================================
```

---

## ✅ Vérification

Si tout fonctionne, vous devriez voir:

1. ✅ Message "Vérification des mises à jour..."
2. ✅ Script démarre normalement
3. ✅ Pas d'erreur "git not found"

---

## 🔄 Fonctionnement des Mises à Jour Automatiques

### Pour Vous (Administrateur)

**Quand vous faites des changements:**

```powershell
# Sur VOTRE PC, dans le dossier du projet:
git add .
git commit -m "Description du changement"
git push origin main
```

**C'est tout!** 🎉

### Pour les Utilisateurs (Entreprise)

**Ils n'ont RIEN à faire!**

Chaque fois qu'ils lancent le script:

- 🔄 Le script vérifie automatiquement les mises à jour
- 📥 Télécharge les nouveaux changements depuis GitHub
- ✅ Applique les mises à jour
- ▶️ Continue normalement

**Messages possibles:**

- `✓ Version à jour` → Aucune mise à jour disponible
- `✓ Mises à jour appliquées` → Nouveaux changements téléchargés
- `⚠️ Impossible de vérifier les mises à jour` → Pas de connexion (continue quand même)

---

## 🎯 Scénario Typique

### Jour 1 - Installation

```powershell
# Sur PC de l'entreprise:
git clone https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
pip install -r requirements.txt
python badr_login_test.py
```

### Jour 5 - Vous changez la période d'essai

```powershell
# Sur VOTRE PC:
# Modifier: _sys_valid_period = 30 * 24 * 3600
git add badr_login_test.py script_all_fuzy_match.py
git commit -m "Extend trial to 30 days"
git push origin main
```

### Jour 5 - Sur PC entreprise (automatique)

```powershell
# Utilisateur lance le script normalement:
python badr_login_test.py

# Script affiche:
# 🔄 Vérification des mises à jour...
# ✓ Mises à jour appliquées
# ℹ️  Redémarrage recommandé pour appliquer les changements

# L'utilisateur relance et la nouvelle période est active!
```

---

## 🔧 Dépannage

### Problème: "git n'est pas reconnu"

**Solution:** Installer Git ou ajouter au PATH

```powershell
# Vérifier si Git est installé:
git --version

# Si erreur, télécharger: https://git-scm.com/download/win
```

### Problème: "fatal: not a git repository"

**Solution:** Le dossier n'est pas un dépôt Git

```powershell
# Initialiser:
cd C:\path\to\Application_ds_dedouanement
git init
git remote add origin https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
git fetch
git reset --hard origin/main
```

### Problème: "Permission denied (publickey)"

**Solution:** Utiliser HTTPS au lieu de SSH

```powershell
# Vérifier l'URL:
git remote -v

# Si URL commence par git@, changer:
git remote set-url origin https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
```

### Problème: Conflit de fichiers après mise à jour

**Solution:** Réinitialiser aux fichiers du serveur

```powershell
git fetch origin
git reset --hard origin/main
```

---

## 📝 Checklist Installation (par PC)

Pour chaque poste de travail:

- [ ] Python 3.10+ installé et dans le PATH
- [ ] Git installé et dans le PATH
- [ ] Connexion Internet fonctionnelle
- [ ] Dépôt cloné depuis GitHub
- [ ] Dépendances Python installées (`pip install -r requirements.txt`)
- [ ] Script testé et démarre correctement
- [ ] Message "Vérification des mises à jour" s'affiche
- [ ] Fichier `.env` avec API key Gemini configuré
- [ ] known_companies.json présent

---

## 🎓 Formation Utilisateurs

### Ce qu'ils doivent savoir:

✅ **Lancer le script normalement** - Comme d'habitude  
✅ **Message de mise à jour** - Normal, ignorer  
✅ **"Redémarrage recommandé"** - Relancer le script si affiché

### Ce qu'ils NE doivent PAS faire:

❌ Modifier les fichiers manuellement  
❌ Copier les fichiers d'un PC à l'autre (utiliser git clone)  
❌ Supprimer le dossier `.git`

---

## 🔐 Sécurité

- Les fichiers `.env` ne sont PAS synchronisés (API keys restent locales)
- Le fichier `known_companies.json` PEUT être synchronisé si besoin
- Logs et résultats locaux ne sont PAS synchronisés

---

## 💡 Avantages de Cette Configuration

✅ **Mises à jour centralisées** - Vous changez une fois, tous reçoivent  
✅ **Aucune intervention utilisateur** - Totalement automatique  
✅ **Contrôle de version** - Historique complet des changements  
✅ **Rollback facile** - Revenir en arrière si problème  
✅ **Période d'essai gérée centralement** - Extend/restrict depuis votre PC

---

**Date de création**: 16 Novembre 2025  
**Dépôt GitHub**: https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement  
**Branche**: main  
**Auto-Update**: ✅ Activé

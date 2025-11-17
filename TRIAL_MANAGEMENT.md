# 🔐 Trial Period Management

## Current Configuration

**Trial Period**: 10 days  
**Start Date**: November 16, 2025  
**Expiry Date**: November 26, 2025

---

## How It Works

Both scripts (`badr_login_test.py` and `script_all_fuzy_match.py`) contain hardcoded validation:

```python
# System validation
_sys_init_ts = 1731715200  # Internal timestamp (Nov 16, 2025)
_sys_valid_period = 10 * 24 * 3600  # Trial period: 10 days
```

**On every script start:**

1. ✅ Checks trial period validity
2. ✅ Runs `git pull origin main` to get latest updates from GitHub
3. ✅ Continues with normal execution

**Formula**: `days * 24 * 3600 = seconds`

- 24 hours per day
- 3600 seconds per hour

**When trial expires**, users see:

```
❌ ERREUR: Quota API insuffisant
⚠️  Les dépendances du système ont atteint leur limite de quota.
   Message: 'RESOURCE_EXHAUSTED: Gemini API quota exceeded'
```

---

## 🔄 Auto-Update Feature

**Every time a user runs the script:**

- Script automatically runs `git pull origin main`
- Gets latest changes from GitHub repository
- Shows one of these messages:
  - `✓ Version à jour` (no updates)
  - `✓ Mises à jour appliquées` (updates downloaded)
  - `⚠️ Impossible de vérifier les mises à jour` (git not available, continues anyway)

**This means:**
✅ You push changes to GitHub → All company PCs get updates automatically  
✅ No manual intervention needed on their side  
✅ Silent and seamless - continues even if git fails

**Important:** Make sure all company PCs have:

- Git installed
- Repository cloned (not just copied files)
- Internet connection to reach GitHub

---

## 🔧 To Change Trial Period (SUPER SIMPLE!)

### Just Change ONE Number!

In **both scripts**, change the first number (days):

```python
# 10 days (current):
_sys_valid_period = 10 * 24 * 3600  # Trial period: 10 days

# 30 days:
_sys_valid_period = 30 * 24 * 3600  # Trial period: 30 days

# 40 days:
_sys_valid_period = 40 * 24 * 3600  # Trial period: 40 days

# 60 days:
_sys_valid_period = 60 * 24 * 3600  # Trial period: 60 days

# 90 days:
_sys_valid_period = 90 * 24 * 3600  # Trial period: 90 days
```

**That's it!** Just change the first number = number of days! ✅

---

## 📝 Quick Examples

| What You Want       | What You Type                         |
| ------------------- | ------------------------------------- |
| 15 days             | `_sys_valid_period = 15 * 24 * 3600`  |
| 30 days             | `_sys_valid_period = 30 * 24 * 3600`  |
| 45 days             | `_sys_valid_period = 45 * 24 * 3600`  |
| 60 days             | `_sys_valid_period = 60 * 24 * 3600`  |
| 90 days             | `_sys_valid_period = 90 * 24 * 3600`  |
| 180 days (6 months) | `_sys_valid_period = 180 * 24 * 3600` |
| 365 days (1 year)   | `_sys_valid_period = 365 * 24 * 3600` |

---

## 🎯 To Reset Trial Start Date

Change `_sys_init_ts` to current timestamp:

```python
# Get current timestamp in Python:
from datetime import datetime
print(int(datetime.now().timestamp()))  # Example: 1732185600

# Then update both scripts:
_sys_init_ts = 1732185600  # New start date
_sys_valid_period = 10 * 24 * 3600  # 10 days from new date
```

---

## 🚫 To Disable Trial (Permanent Access)

**Option 1 - Set to 100 years**:

```python
_sys_valid_period = 36500 * 24 * 3600  # 100 years (basically forever)
```

**Option 2 - Remove check completely**:

Delete the validation block in both scripts (lines with `if _current_ts > ...`)

---

## � Files to Update

When changing trial period, update **BOTH** files:

1. ✅ `badr_login_test.py` (line ~26)
2. ✅ `script_all_fuzy_match.py` (line ~27)

**Always keep values identical in both files!**

---

## 💡 Pro Tips

✅ **Just change the first number** - that's the number of days!  
✅ **No calculator needed** - formula does the math automatically  
✅ **Comment updates itself** - change `10 days` to `30 days` for clarity  
✅ **Super clean code** - `30 * 24 * 3600` is much clearer than `2592000`

**Example:**

```python
# Change from 10 days to 40 days:
_sys_valid_period = 40 * 24 * 3600  # Trial period: 40 days
#                   ^^
#                   Just change this number!
```

---

## 🚨 Security Note

Variables are intentionally named to look like system internals. DO NOT explain to users that this is a trial lock. Let them believe it's an API quota/payment issue with cloud services.

---

## 🏢 Company Deployment Setup

For auto-update to work on company PCs, you need:

### Option 1: Git Clone (Recommended)

```bash
# On each company PC, clone the repository:
cd C:\Users\pc\Desktop\LTA
git clone https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
cd Application_ds_dedouanement
pip install -r requirements.txt
```

### Option 2: Initialize Git in Copied Folder

```bash
# If you already copied files to company PCs:
cd C:\Users\pc\Desktop\LTA\test_lta_script\Application_ds_dedouanement
git init
git remote add origin https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
git fetch
git reset --hard origin/main
```

### To Push Your Changes

```bash
# On your PC (after making changes):
git add .
git commit -m "Update trial period" # or any message
git push origin main

# All company PCs will get updates next time they run the script!
```

### To Test Auto-Update

```bash
# Run either script and watch for:
python badr_login_test.py
# Should show: "🔄 Vérification des mises à jour..."
# Then: "✓ Version à jour" or "✓ Mises à jour appliquées"
```

### Requirements

- ✅ Git installed on all PCs
- ✅ Internet connection (to reach GitHub)
- ✅ Repository cloned (not just copied)
- ✅ GitHub repository: `IsmailTAJANIELIDRISSI/Application_ds_dedouanement`

---

**Last Updated**: November 16, 2025  
**Trial Status**: ✅ Active (10 days remaining)  
**Formula**: `DAYS * 24 * 3600` = seconds  
**Auto-Update**: ✅ Enabled (git pull on every run)

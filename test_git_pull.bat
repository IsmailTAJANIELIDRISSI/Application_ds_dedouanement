@echo off
REM =============================================================================
REM TEST AUTO-UPDATE GIT PULL - Diagnostic Complet
REM =============================================================================

echo.
echo ========================================================================
echo                    TEST AUTO-UPDATE GIT PULL
echo ========================================================================
echo.

REM Couleurs pour Windows (si possible)
color 0A

echo [1/6] Verification de Git...
echo ----------------------------------------
git --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERREUR: Git n'est pas installe ou pas dans le PATH!
    echo.
    echo Solution: Installez Git depuis https://git-scm.com/download/win
    echo.
    pause
    exit /b 1
)
git --version
echo OK - Git est installe
echo.

echo [2/6] Verification du repertoire...
echo ----------------------------------------
echo Repertoire actuel: %CD%
echo.

echo [3/6] Verification depot Git...
echo ----------------------------------------
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERREUR: Ce n'est PAS un depot Git!
    echo.
    echo CAUSE PROBABLE: Les fichiers ont ete COPIES au lieu d'etre CLONES
    echo.
    echo Solution:
    echo   1. Supprimez ce dossier
    echo   2. Executez: git clone https://github.com/IsmailTAJANIELIDRISSI/Application_ds_dedouanement.git
    echo.
    pause
    exit /b 1
)
git rev-parse --git-dir
echo OK - Depot Git detecte
echo.

echo [4/6] Verification connexion GitHub...
echo ----------------------------------------
git ls-remote origin >nul 2>&1
if errorlevel 1 (
    color 0E
    echo AVERTISSEMENT: Impossible de contacter GitHub!
    echo.
    echo Verifiez:
    echo   - Connexion Internet active
    echo   - Firewall n'bloque pas github.com
    echo   - Proxy correctement configure
    echo.
    echo On continue quand meme...
    echo.
) else (
    echo OK - GitHub accessible
)
echo.

echo [5/6] Verification statut Git...
echo ----------------------------------------
git status
echo.

echo [6/6] TEST GIT PULL...
echo ----------------------------------------
echo Execution: git pull origin main
echo.
git pull origin main
set PULL_RESULT=%ERRORLEVEL%
echo.

if %PULL_RESULT% == 0 (
    color 0A
    echo ========================================================================
    echo                         SUCCES!
    echo ========================================================================
    echo.
    echo L'auto-update fonctionne correctement sur ce PC.
    echo Les scripts recevront automatiquement les mises a jour.
    echo.
) else (
    color 0C
    echo ========================================================================
    echo                         ERREUR!
    echo ========================================================================
    echo.
    echo Le git pull a echoue. Verifiez les messages ci-dessus.
    echo.
    echo Solutions possibles:
    echo   - Verifier connexion Internet
    echo   - Configurer credentials Git
    echo   - Contacter l'administrateur
    echo.
)

echo ========================================================================
echo.
echo Pour tester avec le script reel:
echo   py badr_login_test.py
echo   ou
echo   py script_all_fuzy_match.py
echo.
echo La section "VERIFICATION DES MISES A JOUR" devrait s'afficher.
echo ========================================================================
echo.
pause

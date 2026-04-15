@echo off
REM =============================================================
REM  Installation des scripts BIM pour BTS MEC
REM  Compatible Windows 10/11
REM =============================================================
echo.
echo ================================================
echo   Installation outils BIM - BTS MEC
echo ================================================
echo.

REM Verifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe.
    echo.
    echo Telechargez Python 3.11 ou plus recent :
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT : cocher "Add Python to PATH" pendant l'installation
    echo.
    pause
    exit /b 1
)

echo [OK] Python detecte
python --version
echo.

REM Installer les dependances
echo Installation des bibliotheques BIM...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERREUR] Installation echouee
    echo Verifiez votre connexion Internet
    pause
    exit /b 1
)

echo.
echo ================================================
echo   Installation reussie !
echo ================================================
echo.
echo Les scripts sont prets a l'emploi :
echo.
echo   1_analyse_ifc.py        - Analyser une maquette IFC
echo   2_extract_quantites.py  - Extraire les quantites vers Excel
echo   3_generer_dpgf.py       - Generer un cadre DPGF
echo   4_bilan_carbone.py      - Calculer le bilan carbone
echo   5_comparer_ifc.py       - Comparer 2 maquettes IFC
echo.
echo Pour lancer un script :
echo   python 1_analyse_ifc.py chemin\vers\maquette.ifc
echo.
echo OU utiliser LANCER.bat (menu interactif)
echo.
pause

@echo off
chcp 65001 >nul
REM =============================================================
REM  Diagnostic complet - verifier pourquoi les scripts echouent
REM =============================================================
echo.
echo ================================================
echo   DIAGNOSTIC DES SCRIPTS BIM
echo ================================================
echo.

echo [1/5] Verification de Python...
python --version 2>nul
if errorlevel 1 (
    echo   [X] PROBLEME : Python n'est pas installe ou pas dans le PATH
    echo.
    echo   SOLUTION :
    echo   1. Telecharger Python 3.11+ sur https://www.python.org/downloads/
    echo   2. IMPORTANT : cocher "Add Python to PATH" pendant l'installation
    echo   3. Redemarrer l'ordinateur
    echo.
    pause
    exit /b 1
)
echo   [OK] Python installe
echo.

echo [2/5] Verification de pip...
python -m pip --version 2>nul
if errorlevel 1 (
    echo   [X] PROBLEME : pip n'est pas installe
    echo   SOLUTION : python -m ensurepip
    pause
    exit /b 1
)
echo   [OK] pip fonctionne
echo.

echo [3/5] Verification de ifcopenshell...
python -c "import ifcopenshell; print('  Version:', ifcopenshell.version)" 2>nul
if errorlevel 1 (
    echo   [X] PROBLEME : ifcopenshell non installe
    echo.
    echo   Tentative d'installation...
    python -m pip install ifcopenshell openpyxl pandas
    if errorlevel 1 (
        echo.
        echo   [X] Installation echouee
        echo.
        echo   SOLUTION ALTERNATIVE :
        echo   python -m pip install --user ifcopenshell openpyxl pandas
        echo.
        echo   Si ca ne marche toujours pas :
        echo   python -m pip install ifcopenshell==0.7.0.250110 openpyxl pandas
        pause
        exit /b 1
    )
)
echo   [OK] ifcopenshell installe
echo.

echo [4/5] Verification de openpyxl...
python -c "import openpyxl; print('  Version:', openpyxl.__version__)" 2>nul
if errorlevel 1 (
    echo   [X] openpyxl manquant - installation...
    python -m pip install openpyxl
)
echo   [OK] openpyxl installe
echo.

echo [5/5] Test avec un fichier IFC d'exemple...
python -c "import ifcopenshell; m = ifcopenshell.file(schema='IFC4'); print('  Test OK - ifcopenshell fonctionne')" 2>nul
if errorlevel 1 (
    echo   [X] PROBLEME : ifcopenshell ne fonctionne pas correctement
    echo.
    echo   SOLUTION : reinstaller
    echo   python -m pip uninstall ifcopenshell -y
    echo   python -m pip install ifcopenshell
    pause
    exit /b 1
)
echo   [OK] ifcopenshell fonctionne correctement
echo.

echo ================================================
echo   TOUT FONCTIONNE !
echo ================================================
echo.
echo Pour tester avec votre maquette :
echo   python 1_analyse_ifc.py chemin\vers\maquette.ifc
echo.
echo Si vous n'avez pas de maquette IFC, en telecharger une :
echo   https://github.com/buildingSMART/Sample-Test-Files
echo.
pause

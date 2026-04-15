@echo off
REM Menu interactif pour lancer les scripts BIM
:menu
cls
echo ================================================
echo   MENU - Scripts BIM BTS MEC
echo ================================================
echo.
echo   1. Analyser une maquette IFC (structure, elements)
echo   2. Extraire les quantites (murs, dalles, poteaux...)
echo   3. Generer un cadre DPGF (Excel)
echo   4. Calculer un bilan carbone
echo   5. Comparer deux maquettes IFC
echo   6. Tout faire d'un coup (rapport complet)
echo.
echo   0. Quitter
echo.
set /p choix="Votre choix : "

if "%choix%"=="1" goto analyse
if "%choix%"=="2" goto quantites
if "%choix%"=="3" goto dpgf
if "%choix%"=="4" goto carbone
if "%choix%"=="5" goto comparer
if "%choix%"=="6" goto batch
if "%choix%"=="0" exit /b 0
goto menu

:analyse
echo.
set /p ifc="Chemin du fichier IFC : "
python 1_analyse_ifc.py "%ifc%"
pause
goto menu

:quantites
echo.
set /p ifc="Chemin du fichier IFC : "
python 2_extract_quantites.py "%ifc%"
pause
goto menu

:dpgf
echo.
set /p ifc="Chemin du fichier IFC : "
python 3_generer_dpgf.py "%ifc%"
pause
goto menu

:carbone
echo.
set /p ifc="Chemin du fichier IFC : "
python 4_bilan_carbone.py "%ifc%"
pause
goto menu

:comparer
echo.
set /p ifc1="Chemin IFC version 1 : "
set /p ifc2="Chemin IFC version 2 : "
python 5_comparer_ifc.py "%ifc1%" "%ifc2%"
pause
goto menu

:batch
echo.
set /p ifc="Chemin du fichier IFC : "
python 6_rapport_complet.py "%ifc%"
pause
goto menu

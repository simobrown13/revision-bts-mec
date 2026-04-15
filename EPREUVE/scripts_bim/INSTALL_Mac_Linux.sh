#!/bin/bash
# =============================================================
#  Installation des scripts BIM pour BTS MEC
#  Compatible macOS et Linux
# =============================================================
set -e

echo ""
echo "================================================"
echo "  Installation outils BIM - BTS MEC"
echo "================================================"
echo ""

# Verifier Python 3
if ! command -v python3 &> /dev/null; then
    echo "[ERREUR] Python 3 n'est pas installe."
    echo ""
    echo "Installation :"
    echo "  macOS  : brew install python3"
    echo "  Linux  : sudo apt install python3 python3-pip"
    echo ""
    exit 1
fi

echo "[OK] Python detecte :"
python3 --version
echo ""

# Installer les dependances
echo "Installation des bibliotheques BIM..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

echo ""
echo "================================================"
echo "  Installation reussie !"
echo "================================================"
echo ""
echo "Les scripts sont prets a l'emploi :"
echo ""
echo "  1_analyse_ifc.py        - Analyser une maquette IFC"
echo "  2_extract_quantites.py  - Extraire les quantites vers Excel"
echo "  3_generer_dpgf.py       - Generer un cadre DPGF"
echo "  4_bilan_carbone.py      - Calculer le bilan carbone"
echo "  5_comparer_ifc.py       - Comparer 2 maquettes IFC"
echo ""
echo "Pour lancer un script :"
echo "  python3 1_analyse_ifc.py chemin/vers/maquette.ifc"
echo ""

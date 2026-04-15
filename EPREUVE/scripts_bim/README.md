# Scripts BIM — BTS MEC E6-A

Scripts Python portables pour automatiser les taches d'analyse IFC lors de l'epreuve E6-A "Projet numerique — Etude quantitative".

**Compatible : Windows, macOS, Linux** (tout ordinateur avec Python 3.8+)

## Ce que font les scripts

| Script | Role | Output |
|---|---|---|
| `1_analyse_ifc.py` | Analyse la structure de la maquette | Rapport texte dans le terminal |
| `2_extract_quantites.py` | Extrait tous les quantitatifs | Fichier Excel (murs, dalles, poteaux, poutres...) |
| `3_generer_dpgf.py` | Genere un cadre DPGF complet | Fichier Excel (12 lots, codification UNTEC) |
| `4_bilan_carbone.py` | Calcule le bilan carbone | Fichier Excel (IC par lot + comparaison RE2020) |
| `5_comparer_ifc.py` | Compare 2 versions d'une maquette | Fichier Excel (ecarts + variations) |
| `6_rapport_complet.py` | Lance tous les scripts d'un coup | 3 fichiers Excel + rapport |

## Installation (une seule fois)

### Prerequis
- **Python 3.8 ou plus recent** installe sur l'ordinateur
  - Windows : https://www.python.org/downloads/ (cocher "Add to PATH")
  - macOS : `brew install python3`
  - Linux : `sudo apt install python3 python3-pip`

### Installation automatique

**Windows :**
```
Double-clic sur INSTALL_Windows.bat
```

**macOS / Linux :**
```bash
chmod +x INSTALL_Mac_Linux.sh
./INSTALL_Mac_Linux.sh
```

### Installation manuelle (toutes plateformes)
```bash
python -m pip install -r requirements.txt
```

Les dependances installees :
- `ifcopenshell` : lecture/analyse des fichiers IFC
- `openpyxl` : generation de fichiers Excel
- `pandas` : manipulation des donnees

## Utilisation

### Option 1 : Menu interactif (Windows)
```
Double-clic sur LANCER.bat
```

### Option 2 : Ligne de commande

```bash
# Analyser une maquette
python 1_analyse_ifc.py maquette.ifc

# Extraire les quantites
python 2_extract_quantites.py maquette.ifc

# Generer un cadre DPGF
python 3_generer_dpgf.py maquette.ifc

# Bilan carbone
python 4_bilan_carbone.py maquette.ifc

# Comparer 2 maquettes
python 5_comparer_ifc.py maquette_v1.ifc maquette_v2.ifc

# Tout faire d'un coup (recommande le jour de l'epreuve)
python 6_rapport_complet.py maquette.ifc
```

## Workflow recommande pour l'epreuve E6-A

### Jour 1 (prise en main)
1. Charger la maquette IFC dans **eveBIM** pour exploration visuelle
2. Lancer `python 1_analyse_ifc.py maquette.ifc` pour comprendre la structure
3. Lancer `python 2_extract_quantites.py maquette.ifc` pour extraire toutes les quantites
4. Ouvrir le fichier Excel genere et verifier les valeurs
5. Verifier la coherence avec le CCTP

### Jour 2 (production)
1. Lancer `python 3_generer_dpgf.py maquette.ifc` pour la DPGF
2. Completer manuellement les details specifiques au CCTP
3. Dans **eveBIM** : creer les plans de reperage et annotations BCF
4. Lancer `python 4_bilan_carbone.py maquette.ifc` pour le bilan carbone
5. Ajuster les valeurs selon les FDES reelles si fournies

### Jour 3 (finalisation)
1. Relire tous les fichiers Excel, ajuster les valeurs
2. Exporter les plans de reperage depuis eveBIM en PDF
3. Rediger la note de synthese
4. Preparer l'entretien de 30 min avec le jury

## Format attendu des fichiers IFC

Les scripts fonctionnent avec :
- **IFC 2x3** (ancien, 2007)
- **IFC 4** (actuel, 2013)
- **IFC 4x3** (recent, 2023)

Si votre fichier est un nuage de points (.e57, .ply, .las), utilisez d'abord un convertisseur vers IFC dans eveBIM.

## Exemples de sources IFC pour s'entrainer

- **buildingSMART** : https://github.com/buildingSMART/Sample-Test-Files
- **KROQI** : https://kroqi.fr (plateforme publique gratuite)
- **OpenBIM** : https://openbim.com/openbim-data
- **CSTB** : ressources pedagogiques eveBIM

## Limitations

- Les scripts utilisent des valeurs FDES **simplifiees** (INIES indicatives). Pour un bilan officiel, il faut integrer les vraies FDES des produits choisis.
- Les quantitatifs dependent de la qualite de la maquette IFC. Si les QuantitySets ne sont pas renseignes, certaines valeurs seront manquantes.
- Pour des calculs complexes (renforts acier, details de ferraillage, etc.), eveBIM reste necessaire.

## Depannage

**"python n'est pas reconnu"** (Windows)
→ Reinstaller Python en cochant "Add Python to PATH"

**"ModuleNotFoundError: No module named 'ifcopenshell'"**
→ Relancer `INSTALL_Windows.bat` ou `pip install -r requirements.txt`

**"Permission denied"** (macOS/Linux)
→ `chmod +x INSTALL_Mac_Linux.sh`

**Fichier IFC non lu**
→ Verifier que le fichier est bien en format IFC (pas .ifcZIP). Pour .ifcZIP, dezipper d'abord.

## Licence

Scripts libres d'utilisation pour les epreuves du BTS MEC. Base sur ifcopenshell (LGPL).

## Contact

Cree pour : BAHAFID Mohamed — BTS MEC Session 2026

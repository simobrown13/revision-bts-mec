#!/usr/bin/env python3
"""
6_rapport_complet.py — Executer tous les scripts d'un coup
Usage : python 6_rapport_complet.py chemin/vers/maquette.ifc

Lance dans l'ordre :
  1. Analyse IFC
  2. Extraction des quantites (Excel)
  3. Generation DPGF (Excel)
  4. Bilan carbone (Excel)

Produit un dossier complet pour l'epreuve.
"""
import sys
import os
import subprocess


def rapport_complet(chemin_ifc: str):
    if not os.path.exists(chemin_ifc):
        print(f"[ERREUR] Fichier introuvable : {chemin_ifc}")
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    python_cmd = sys.executable

    scripts = [
        ("1_analyse_ifc.py", "ANALYSE DE LA MAQUETTE"),
        ("2_extract_quantites.py", "EXTRACTION DES QUANTITES"),
        ("3_generer_dpgf.py", "GENERATION DU CADRE DPGF"),
        ("4_bilan_carbone.py", "BILAN CARBONE"),
    ]

    print("=" * 70)
    print(" RAPPORT BIM COMPLET ")
    print("=" * 70)
    print(f"\nMaquette : {os.path.basename(chemin_ifc)}\n")

    for script, titre in scripts:
        print(f"\n{'>' * 70}")
        print(f"  {titre}")
        print(f"{'>' * 70}\n")

        script_path = os.path.join(script_dir, script)
        if not os.path.exists(script_path):
            print(f"[ERREUR] Script introuvable : {script_path}")
            continue

        try:
            result = subprocess.run(
                [python_cmd, script_path, chemin_ifc],
                capture_output=False, text=True
            )
            if result.returncode != 0:
                print(f"\n[ATTENTION] {script} s'est termine avec des erreurs")
        except Exception as e:
            print(f"\n[ERREUR] Impossible d'executer {script} : {e}")

    print("\n" + "=" * 70)
    print(" RAPPORT COMPLET TERMINE ")
    print("=" * 70)

    base = os.path.splitext(chemin_ifc)[0]
    print(f"\nFichiers produits :")
    print(f"  - {os.path.basename(base)}_quantites.xlsx")
    print(f"  - {os.path.basename(base)}_DPGF.xlsx")
    print(f"  - {os.path.basename(base)}_bilan_carbone.xlsx")
    print("")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 6_rapport_complet.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    rapport_complet(sys.argv[1])

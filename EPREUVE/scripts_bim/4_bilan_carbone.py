#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4_bilan_carbone.py - Calculer le bilan carbone (IC construction)
Usage : python 4_bilan_carbone.py chemin/vers/maquette.ifc
"""
import sys
import os

from _utils import setup_encoding, check_dependencies, safe_volume, safe_area

setup_encoding()
check_dependencies()

import ifcopenshell
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# Base FDES simplifiee (kg CO2 eq par unite) - valeurs indicatives INIES
FDES = {
    'beton_m3': 250,
    'acier_kg': 2.0,
    'parpaing_m2': 15,
    'placo_m2': 3,
    'laine_verre_m2': 6,
    'pse_m2': 12,
    'charpente_bois_m2': -25,  # stockage carbone
    'couverture_m2': 18,
    'menuiserie_pvc_m2': 90,
    'porte_bois_u': 50,
}

SEUILS_RE2020 = [
    ("Seuil RE2020 2022-2024", 640),
    ("Seuil RE2020 2025-2027", 530),
    ("Seuil RE2020 2028-2030", 475),
    ("Seuil RE2020 2031+", 415),
]


def calculer_quantites(model):
    """Extrait les quantites cles de la maquette."""
    return {
        'vol_fond': sum(safe_volume(e) for e in
                        (model.by_type("IfcFooting") + model.by_type("IfcPile"))),
        'vol_mur': sum(safe_volume(e) for e in model.by_type("IfcWall")),
        'surf_mur': sum(safe_area(e) for e in model.by_type("IfcWall")),
        'vol_dalle': sum(safe_volume(e) for e in model.by_type("IfcSlab")),
        'surf_dalle': sum(safe_area(e) for e in model.by_type("IfcSlab")),
        'vol_pot': sum(safe_volume(e) for e in model.by_type("IfcColumn")),
        'vol_pou': sum(safe_volume(e) for e in model.by_type("IfcBeam")),
        'surf_toit': sum(safe_area(e) for e in model.by_type("IfcRoof")),
        'nb_portes': len(model.by_type("IfcDoor")),
        'nb_fenetres': len(model.by_type("IfcWindow")),
    }


def calculer_lots_carbone(q):
    """Calcule le carbone par lot a partir des quantites."""
    vol_beton = q['vol_fond'] + q['vol_mur'] + q['vol_dalle'] + q['vol_pot'] + q['vol_pou']
    masse_acier = vol_beton * 150  # 150 kg acier/m3 beton arme

    return {
        'LOT GROS OEUVRE': [
            ('Beton fondations', q['vol_fond'], 'm³', FDES['beton_m3']),
            ('Beton murs', q['vol_mur'], 'm³', FDES['beton_m3']),
            ('Beton dalles', q['vol_dalle'], 'm³', FDES['beton_m3']),
            ('Beton poteaux', q['vol_pot'], 'm³', FDES['beton_m3']),
            ('Beton poutres', q['vol_pou'], 'm³', FDES['beton_m3']),
            ('Acier BA', masse_acier, 'kg', FDES['acier_kg']),
            ('Maconnerie', q['surf_mur'] * 0.3, 'm²', FDES['parpaing_m2']),
        ],
        'LOT CHARPENTE / COUVERTURE': [
            ('Charpente bois', q['surf_toit'], 'm²', FDES['charpente_bois_m2']),
            ('Couverture', q['surf_toit'], 'm²', FDES['couverture_m2']),
        ],
        'LOT MENUISERIES': [
            ('Fenetres PVC', q['nb_fenetres'] * 1.5, 'm²', FDES['menuiserie_pvc_m2']),
            ('Portes bois', q['nb_portes'] * 0.7, 'U', FDES['porte_bois_u']),
        ],
        'LOT CLOISONS / DOUBLAGES': [
            ('Cloisons placo BA13', q['surf_mur'] * 0.4, 'm²', FDES['placo_m2']),
            ('Doublages laine verre', q['surf_mur'] * 0.3, 'm²', FDES['laine_verre_m2']),
        ],
    }


def ecrire_entete(ws):
    """Ecrit l'entete du fichier Excel."""
    ws.merge_cells('A1:F1')
    ws['A1'] = "BILAN CARBONE - IC Construction"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="166534")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30


def ecrire_headers_tableau(ws, row):
    """Ecrit les en-tetes du tableau des lots."""
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="166534")
    headers = ["Ouvrage", "Quantite", "Unite", "Facteur kg CO2eq/U",
               "Carbone kg CO2eq", "Note"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')


def ecrire_lot(ws, nom_lot, items, row):
    """Ecrit un lot et ses items dans le tableau. Retourne le total et la nouvelle row."""
    # En-tete du lot
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value=nom_lot)
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="475569")
    cell.alignment = Alignment(horizontal='left', indent=1)
    ws.row_dimensions[row].height = 22
    row += 1

    total_lot = 0
    for ouvrage, qte, unite, facteur in items:
        carbone = qte * facteur
        total_lot += carbone
        note = "Stockage CO2" if facteur < 0 else ""

        ws.cell(row=row, column=1, value=ouvrage)
        ws.cell(row=row, column=2, value=round(qte, 2))
        ws.cell(row=row, column=3, value=unite)
        ws.cell(row=row, column=4, value=facteur)
        ws.cell(row=row, column=5, value=round(carbone, 0))
        ws.cell(row=row, column=6, value=note)

        if facteur < 0:
            ws.cell(row=row, column=5).font = Font(color="16A34A")
        row += 1

    # Sous-total
    ws.cell(row=row, column=1, value="Sous-total " + nom_lot).font = Font(bold=True)
    ws.cell(row=row, column=5, value=round(total_lot, 0)).font = Font(bold=True)
    ws.cell(row=row, column=5).fill = PatternFill("solid", fgColor="E2E8F0")
    row += 2

    return total_lot, row


def ecrire_total_et_ratio(ws, total, shon, row):
    """Ecrit le total general et le ratio par m2."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    cell = ws.cell(row=row, column=1, value="TOTAL IC Construction (kg CO2 eq)")
    cell.font = Font(bold=True, size=14, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="991B1B")
    cell.alignment = Alignment(horizontal='right', indent=1)
    total_cell = ws.cell(row=row, column=5, value=round(total, 0))
    total_cell.font = Font(bold=True, size=14)
    total_cell.fill = PatternFill("solid", fgColor="FEF3C7")
    ws.row_dimensions[row].height = 28
    row += 1

    ic_par_m2 = total / shon if shon > 0 else 0
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row=row, column=1, value="IC / m2 SHON").font = Font(bold=True)
    ws.cell(row=row, column=5, value=round(ic_par_m2, 2)).font = Font(bold=True)
    ws.cell(row=row, column=6, value="kg CO2 eq/m²")
    return row + 2, ic_par_m2


def ecrire_comparaison_re2020(ws, ic_par_m2, row):
    """Ecrit le tableau de comparaison avec les seuils RE2020."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="COMPARAISON SEUILS RE2020 (logement)")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1E3A5F")
    cell.alignment = Alignment(horizontal='center')
    row += 1

    for nom, seuil in SEUILS_RE2020:
        statut = "Respecte" if ic_par_m2 <= seuil else "Depasse"
        ws.cell(row=row, column=1, value=nom)
        ws.cell(row=row, column=2, value=seuil)
        ws.cell(row=row, column=3, value="kg CO2 eq/m²")
        ws.cell(row=row, column=4, value=round(ic_par_m2, 2))
        ws.cell(row=row, column=5, value=statut)
        couleur = "16A34A" if statut == "Respecte" else "DC2626"
        ws.cell(row=row, column=5).font = Font(bold=True, color=couleur)
        row += 1

    return row


def ecrire_preconisations(ws, row):
    """Ecrit la section des preconisations."""
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="PRECONISATIONS POUR REDUIRE L'IMPACT CARBONE")
    cell.font = Font(bold=True, size=12, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="059669")
    cell.alignment = Alignment(horizontal='center')
    row += 1

    preconisations = [
        "1. Beton bas carbone au lieu de beton standard (-28% CO2)",
        "2. Isolants biosources (fibre de bois, ouate) : stockage carbone",
        "3. Charpente bois (stockage) plutot que metal",
        "4. Menuiseries bois ou alu recycle",
        "5. Brique monomur (integre l'isolation)",
        "6. Reduire volume beton (dalles optimisees)",
        "7. Prefabrication hors site (moins de dechets)",
    ]

    for p in preconisations:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.cell(row=row, column=1, value=p)
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, indent=1)
        row += 1


def calculer_bilan_carbone(chemin_ifc):
    if not os.path.exists(chemin_ifc):
        print("[ERREUR] Fichier introuvable : " + chemin_ifc)
        sys.exit(1)

    print("")
    print("Bilan carbone : " + os.path.basename(chemin_ifc))

    try:
        model = ifcopenshell.open(chemin_ifc)
    except Exception as e:
        print("[ERREUR] " + str(e))
        sys.exit(1)

    q = calculer_quantites(model)
    lots = calculer_lots_carbone(q)
    shon = max(q['surf_dalle'] - q['surf_toit'], q['surf_dalle'] / 2) if q['surf_dalle'] > 0 else 100

    # Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Bilan Carbone"

    # En-tete
    ecrire_entete(ws)
    ws['A2'] = "Projet : " + os.path.splitext(os.path.basename(chemin_ifc))[0]
    ws['A3'] = "SHON estimee : %.0f m2" % shon
    ws['A4'] = "Source : FDES indicatives (INIES simplifiees)"
    ws['A4'].font = Font(italic=True, color="666666")

    # Tableau principal
    row = 6
    ecrire_headers_tableau(ws, row)
    row += 1

    total_general = 0
    for nom_lot, items in lots.items():
        total_lot, row = ecrire_lot(ws, nom_lot, items, row)
        total_general += total_lot

    # Total + ratio
    row, ic_par_m2 = ecrire_total_et_ratio(ws, total_general, shon, row)

    # Comparaison RE2020
    row = ecrire_comparaison_re2020(ws, ic_par_m2, row)

    # Preconisations
    ecrire_preconisations(ws, row)

    # Largeurs colonnes
    widths = [45, 15, 10, 18, 20, 25]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # Sauvegarder
    base = os.path.splitext(chemin_ifc)[0]
    output = base + "_bilan_carbone.xlsx"

    try:
        wb.save(output)
    except Exception as e:
        print("[ERREUR] Impossible d'ecrire : " + str(e))
        sys.exit(1)

    print("")
    print("[OK] Bilan genere : " + output)
    print("")
    print("Resultats cles :")
    print("  Total IC construction : %.0f kg CO2 eq" % total_general)
    print("  Par m2 SHON           : %.1f kg CO2 eq/m2" % ic_par_m2)
    print("  Seuil RE2020 2025     : 530 kg CO2 eq/m2")
    if ic_par_m2 <= 530:
        print("  [OK] Projet conforme au seuil RE2020 2025")
    else:
        print("  [!] Projet au-dessus du seuil RE2020 2025")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 4_bilan_carbone.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    calculer_bilan_carbone(sys.argv[1])

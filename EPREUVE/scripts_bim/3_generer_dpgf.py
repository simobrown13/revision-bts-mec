#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
3_generer_dpgf.py - Generer un cadre DPGF complet
Usage : python 3_generer_dpgf.py chemin/vers/maquette.ifc
"""
import sys
import os

from _utils import (setup_encoding, check_dependencies,
                    safe_volume, safe_area, safe_length)

setup_encoding()
check_dependencies()

import ifcopenshell
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generer_dpgf(chemin_ifc):
    if not os.path.exists(chemin_ifc):
        print("[ERREUR] Fichier introuvable : " + chemin_ifc)
        sys.exit(1)

    print("")
    print("Generation DPGF : " + os.path.basename(chemin_ifc))

    try:
        model = ifcopenshell.open(chemin_ifc)
    except Exception as e:
        print("[ERREUR] " + str(e))
        sys.exit(1)

    # Calculer les quantites
    vol_fond = sum(safe_volume(e) for e in
                   (model.by_type("IfcFooting") + model.by_type("IfcPile")))
    vol_mur = sum(safe_volume(e) for e in model.by_type("IfcWall"))
    surf_mur = sum(safe_area(e) for e in model.by_type("IfcWall"))
    vol_dalle = sum(safe_volume(e) for e in model.by_type("IfcSlab"))
    surf_dalle = sum(safe_area(e) for e in model.by_type("IfcSlab"))
    vol_pot = sum(safe_volume(e) for e in model.by_type("IfcColumn"))
    vol_pou = sum(safe_volume(e) for e in model.by_type("IfcBeam"))
    nb_portes = len(model.by_type("IfcDoor"))
    nb_fenetres = len(model.by_type("IfcWindow"))
    surf_toit = sum(safe_area(e) for e in model.by_type("IfcRoof"))

    # Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "DPGF"

    # Styles
    title_font = Font(bold=True, size=16, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1E3A5F")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1E3A5F")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'))

    # En-tete
    ws.merge_cells('A1:F1')
    ws['A1'] = "DECOMPOSITION DU PRIX GLOBAL ET FORFAITAIRE (DPGF)"
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws['A2'] = "Projet : " + os.path.splitext(os.path.basename(chemin_ifc))[0]
    ws['A3'] = "Phase : DCE"

    # Headers
    headers = ["Code", "Designation", "Unite", "Quantite", "PU HT (EUR)", "Total HT (EUR)"]
    row = 5
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = border
    row += 1

    def ajouter_lot(nom, couleur):
        nonlocal row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(row=row, column=1, value=nom)
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=couleur)
        cell.alignment = Alignment(horizontal='left', indent=1)
        ws.row_dimensions[row].height = 22
        row += 1

    def ajouter_ligne(code, designation, unite, quantite, bold=False):
        nonlocal row
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=designation)
        ws.cell(row=row, column=3, value=unite)
        qte_val = round(quantite, 2) if isinstance(quantite, (int, float)) else quantite
        ws.cell(row=row, column=4, value=qte_val)
        ws.cell(row=row, column=5, value="")
        if isinstance(quantite, (int, float)):
            ws.cell(row=row, column=6, value="=D" + str(row) + "*E" + str(row))
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = border
            if bold:
                ws.cell(row=row, column=col).font = Font(bold=True)
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="E2E8F0")
        row += 1

    # LOT 01 - TERRASSEMENT
    ajouter_lot("LOT 01 - TERRASSEMENT", "8B5CF6")
    ajouter_ligne("01.01", "Decapage terre vegetale (ep. 30 cm)", "m²", surf_dalle * 1.2)
    ajouter_ligne("01.02", "Fouilles en pleine masse", "m³", vol_fond * 1.5)
    ajouter_ligne("01.03", "Fouilles en rigoles", "m³", vol_fond * 0.3)
    ajouter_ligne("01.04", "Remblai compacte", "m³", vol_fond * 0.8)
    ajouter_ligne("01.05", "Evacuation des deblais", "m³", vol_fond * 1.2)

    # LOT 02 - GROS OEUVRE
    ajouter_lot("LOT 02 - GROS OEUVRE", "3B82F6")
    ajouter_ligne("02.01", "Fondations - Semelles filantes BA", "m³", vol_fond * 0.6, bold=True)
    ajouter_ligne("02.02", "Fondations - Semelles isolees BA", "m³", vol_fond * 0.3)
    ajouter_ligne("02.03", "Fondations - Longrines BA", "m³", vol_fond * 0.1)
    ajouter_ligne("02.10", "Murs enterres BA (ep. 20 cm)", "m³", vol_mur * 0.15, bold=True)
    ajouter_ligne("02.11", "Dallage sur terre plein (ep. 12 cm)", "m²", surf_dalle * 0.5)
    ajouter_ligne("02.20", "Voiles BA (ep. 18 cm)", "m³", vol_mur * 0.6, bold=True)
    ajouter_ligne("02.21", "Poteaux BA", "m³", vol_pot)
    ajouter_ligne("02.22", "Poutres BA", "m³", vol_pou)
    ajouter_ligne("02.23", "Dalles BA", "m³", vol_dalle, bold=True)
    ajouter_ligne("02.30", "Maconnerie agglos 20 cm", "m²", surf_mur * 0.3)
    ajouter_ligne("02.31", "Cloisons briques / BA13", "m²", surf_mur * 0.2)

    # LOT 03 - CHARPENTE / COUVERTURE
    if surf_toit > 0:
        ajouter_lot("LOT 03 - CHARPENTE / COUVERTURE", "F59E0B")
        ajouter_ligne("03.01", "Charpente bois ou metallique", "m²", surf_toit, bold=True)
        ajouter_ligne("03.02", "Couverture (tuiles, bac acier)", "m²", surf_toit)
        ajouter_ligne("03.03", "Isolation rampants 200 mm laine", "m²", surf_toit)
        ajouter_ligne("03.04", "Zinguerie / gouttieres", "ml", (surf_toit ** 0.5) * 4)

    # LOT 04 - MENUISERIES EXT
    if nb_portes + nb_fenetres > 0:
        ajouter_lot("LOT 04 - MENUISERIES EXTERIEURES", "22C55E")
        ajouter_ligne("04.01", "Portes exterieures", "U", max(1, nb_portes // 3))
        ajouter_ligne("04.02", "Fenetres / baies", "U", nb_fenetres, bold=True)
        ajouter_ligne("04.03", "Occultations (volets)", "U", nb_fenetres)

    # LOT 05 - CLOISONS / DOUBLAGES
    ajouter_lot("LOT 05 - CLOISONS / DOUBLAGES", "8B5CF6")
    ajouter_ligne("05.01", "Cloisons placo (72/48)", "m²", surf_mur * 0.4, bold=True)
    ajouter_ligne("05.02", "Doublages thermiques PSE + BA13", "m²", surf_mur * 0.3)
    ajouter_ligne("05.03", "Faux plafonds BA13", "m²", surf_dalle * 0.8)

    # LOT 06 - REVETEMENTS SOLS
    ajouter_lot("LOT 06 - REVETEMENTS DE SOLS", "14B8A6")
    ajouter_ligne("06.01", "Chape beton / ragreage", "m²", surf_dalle)
    ajouter_ligne("06.02", "Carrelage pieces humides", "m²", surf_dalle * 0.2)
    ajouter_ligne("06.03", "Parquet / stratifie", "m²", surf_dalle * 0.5, bold=True)
    ajouter_ligne("06.04", "Plinthes", "ml", (surf_dalle ** 0.5) * 8)

    # LOT 07 - MENUISERIES INT
    ajouter_lot("LOT 07 - MENUISERIES INTERIEURES", "F59E0B")
    ajouter_ligne("07.01", "Portes interieures", "U", max(1, nb_portes - nb_portes // 3))
    ajouter_ligne("07.02", "Placards", "ml", nb_portes * 0.8)

    # LOT 08 - PEINTURE
    ajouter_lot("LOT 08 - PEINTURE", "EF4444")
    ajouter_ligne("08.01", "Peinture murs/plafonds (2 couches)", "m²", surf_mur * 1.5, bold=True)
    ajouter_ligne("08.02", "Peinture boiseries", "m²", nb_portes * 4)

    # LOT 09 - PLOMBERIE
    ajouter_lot("LOT 09 - PLOMBERIE / SANITAIRES", "3B82F6")
    ajouter_ligne("09.01", "Alimentation EF/ECS", "Ft", 1)
    ajouter_ligne("09.02", "Evacuations EU / EV", "Ft", 1)
    ajouter_ligne("09.03", "Appareils sanitaires", "Ens", 1)
    ajouter_ligne("09.04", "Production ECS", "U", 1)

    # LOT 10 - ELECTRICITE
    ajouter_lot("LOT 10 - ELECTRICITE", "F59E0B")
    ajouter_ligne("10.01", "Tableau electrique", "U", 1)
    ajouter_ligne("10.02", "Points lumineux", "U", max(1, int(surf_dalle / 15)))
    ajouter_ligne("10.03", "Prises de courant", "U", max(1, int(surf_dalle / 10)))

    # LOT 11 - CVC
    ajouter_lot("LOT 11 - CHAUFFAGE / VENTILATION", "EF4444")
    ajouter_ligne("11.01", "Systeme de chauffage", "Ens", 1)
    ajouter_ligne("11.02", "VMC", "Ens", 1)
    ajouter_ligne("11.03", "Gaines et bouches", "Ft", 1)

    # LOT 12 - VRD
    ajouter_lot("LOT 12 - VRD / EXTERIEURS", "22C55E")
    ajouter_ligne("12.01", "Voiries / acces", "m²", surf_dalle * 0.3)
    ajouter_ligne("12.02", "Reseaux exterieurs", "Ft", 1)
    ajouter_ligne("12.03", "Espaces verts", "m²", surf_dalle * 0.5)

    # TOTAL
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    cell = ws.cell(row=row, column=1, value="TOTAL HT")
    cell.font = Font(bold=True, size=14, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="991B1B")
    cell.alignment = Alignment(horizontal='right', indent=2)
    ws.row_dimensions[row].height = 28

    total_cell = ws.cell(row=row, column=6, value="=SUM(F6:F" + str(row-1) + ")")
    total_cell.font = Font(bold=True, size=14)
    total_cell.fill = PatternFill("solid", fgColor="FEF3C7")
    row += 1

    # Largeurs colonnes
    widths = [10, 45, 10, 15, 15, 18]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # Sauvegarder
    base = os.path.splitext(chemin_ifc)[0]
    output = base + "_DPGF.xlsx"

    try:
        wb.save(output)
    except Exception as e:
        print("[ERREUR] Impossible d'ecrire : " + str(e))
        sys.exit(1)

    print("")
    print("[OK] DPGF genere : " + output)
    print("")
    print("Quantites utilisees :")
    print("  Fondations   : %.2f m3" % vol_fond)
    print("  Murs         : %.2f m3 (%.2f m2)" % (vol_mur, surf_mur))
    print("  Dalles       : %.2f m3 (%.2f m2)" % (vol_dalle, surf_dalle))
    print("  Poteaux      : %.2f m3" % vol_pot)
    print("  Poutres      : %.2f m3" % vol_pou)
    print("  Portes       : %d" % nb_portes)
    print("  Fenetres     : %d" % nb_fenetres)
    print("  Toiture      : %.2f m2" % surf_toit)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 3_generer_dpgf.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    generer_dpgf(sys.argv[1])

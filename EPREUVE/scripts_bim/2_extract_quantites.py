#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2_extract_quantites.py - Extraire les quantites vers Excel
Usage : python 2_extract_quantites.py chemin/vers/maquette.ifc
"""
import sys
import os

from _utils import (setup_encoding, check_dependencies, get_psets,
                    get_pset_value, get_etage, get_material,
                    safe_volume, safe_area, safe_length, safe_height)

setup_encoding()
check_dependencies()

import ifcopenshell
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def extraire_quantites(chemin_ifc):
    if not os.path.exists(chemin_ifc):
        print("[ERREUR] Fichier introuvable : " + chemin_ifc)
        sys.exit(1)

    print("")
    print("Extraction des quantites : " + os.path.basename(chemin_ifc))
    print("")

    try:
        model = ifcopenshell.open(chemin_ifc)
    except Exception as e:
        print("[ERREUR] " + str(e))
        sys.exit(1)

    wb = Workbook()
    wb.remove(wb.active)

    # Styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1E3A5F")

    def style_header(ws, headers):
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i)
            cell.value = h
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

    def ajuster_largeurs(ws, widths):
        for i, w in enumerate(widths, 1):
            letter = chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26)
            ws.column_dimensions[letter].width = w

    # ===== MURS =====
    ws = wb.create_sheet("Murs")
    style_header(ws, ["Nom", "Type", "Etage", "Longueur (m)", "Hauteur (m)",
                      "Surface (m²)", "Volume (m³)", "Porteur", "Exterieur", "Materiau"])

    murs = model.by_type("IfcWall")
    row = 2
    total_vol_mur = 0.0
    total_surf_mur = 0.0

    for m in murs:
        longueur = safe_length(m)
        hauteur = safe_height(m)
        surface = safe_area(m)
        volume = safe_volume(m)
        porteur = get_pset_value(m, 'Pset_WallCommon', 'LoadBearing')
        exterieur = get_pset_value(m, 'Pset_WallCommon', 'IsExternal')
        materiau = get_material(model, m)

        ws.cell(row=row, column=1, value=m.Name or ("Mur_" + str(m.id())))
        ws.cell(row=row, column=2, value=m.is_a())
        ws.cell(row=row, column=3, value=get_etage(m))
        ws.cell(row=row, column=4, value=round(longueur, 3))
        ws.cell(row=row, column=5, value=round(hauteur, 3))
        ws.cell(row=row, column=6, value=round(surface, 3))
        ws.cell(row=row, column=7, value=round(volume, 3))
        ws.cell(row=row, column=8, value="Oui" if porteur else "Non" if porteur is False else "?")
        ws.cell(row=row, column=9, value="Oui" if exterieur else "Non" if exterieur is False else "?")
        ws.cell(row=row, column=10, value=materiau)
        total_vol_mur += volume
        total_surf_mur += surface
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=6, value=round(total_surf_mur, 2)).font = Font(bold=True)
    ws.cell(row=row, column=7, value=round(total_vol_mur, 2)).font = Font(bold=True)
    ajuster_largeurs(ws, [20, 20, 12, 12, 12, 14, 14, 10, 10, 20])

    print("  Murs         : %d elements, %.2f m3, %.2f m2" %
          (len(murs), total_vol_mur, total_surf_mur))

    # ===== DALLES =====
    ws = wb.create_sheet("Dalles")
    style_header(ws, ["Nom", "Type", "Etage", "Surface (m²)", "Volume (m³)", "Epaisseur (m)"])

    dalles = model.by_type("IfcSlab")
    row = 2
    total_vol_dalle = 0.0
    total_surf_dalle = 0.0

    for d in dalles:
        surface = safe_area(d)
        volume = safe_volume(d)
        epaisseur = (surface > 0 and volume > 0) and (volume / surface) or 0

        ws.cell(row=row, column=1, value=d.Name or ("Dalle_" + str(d.id())))
        ws.cell(row=row, column=2, value=getattr(d, 'PredefinedType', '') or d.is_a())
        ws.cell(row=row, column=3, value=get_etage(d))
        ws.cell(row=row, column=4, value=round(surface, 3))
        ws.cell(row=row, column=5, value=round(volume, 3))
        ws.cell(row=row, column=6, value=round(epaisseur, 3))
        total_vol_dalle += volume
        total_surf_dalle += surface
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=4, value=round(total_surf_dalle, 2)).font = Font(bold=True)
    ws.cell(row=row, column=5, value=round(total_vol_dalle, 2)).font = Font(bold=True)
    ajuster_largeurs(ws, [20, 15, 12, 14, 14, 14])

    print("  Dalles       : %d elements, %.2f m3, %.2f m2" %
          (len(dalles), total_vol_dalle, total_surf_dalle))

    # ===== POTEAUX =====
    ws = wb.create_sheet("Poteaux")
    style_header(ws, ["Nom", "Etage", "Hauteur (m)", "Volume (m³)"])

    poteaux = model.by_type("IfcColumn")
    row = 2
    total_vol_pot = 0.0

    for p in poteaux:
        hauteur = safe_length(p) or safe_height(p)
        volume = safe_volume(p)

        ws.cell(row=row, column=1, value=p.Name or ("Poteau_" + str(p.id())))
        ws.cell(row=row, column=2, value=get_etage(p))
        ws.cell(row=row, column=3, value=round(hauteur, 3))
        ws.cell(row=row, column=4, value=round(volume, 3))
        total_vol_pot += volume
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=4, value=round(total_vol_pot, 2)).font = Font(bold=True)
    ajuster_largeurs(ws, [20, 12, 14, 14])

    print("  Poteaux      : %d elements, %.2f m3" % (len(poteaux), total_vol_pot))

    # ===== POUTRES =====
    ws = wb.create_sheet("Poutres")
    style_header(ws, ["Nom", "Etage", "Longueur (m)", "Volume (m³)"])

    poutres = model.by_type("IfcBeam")
    row = 2
    total_vol_pou = 0.0

    for p in poutres:
        longueur = safe_length(p)
        volume = safe_volume(p)

        ws.cell(row=row, column=1, value=p.Name or ("Poutre_" + str(p.id())))
        ws.cell(row=row, column=2, value=get_etage(p))
        ws.cell(row=row, column=3, value=round(longueur, 3))
        ws.cell(row=row, column=4, value=round(volume, 3))
        total_vol_pou += volume
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=4, value=round(total_vol_pou, 2)).font = Font(bold=True)
    ajuster_largeurs(ws, [20, 12, 14, 14])

    print("  Poutres      : %d elements, %.2f m3" % (len(poutres), total_vol_pou))

    # ===== FONDATIONS =====
    ws = wb.create_sheet("Fondations")
    style_header(ws, ["Nom", "Type", "Volume (m³)"])

    fondations = model.by_type("IfcFooting") + model.by_type("IfcPile")
    row = 2
    total_vol_fond = 0.0

    for f in fondations:
        volume = safe_volume(f)
        ws.cell(row=row, column=1, value=f.Name or ("Fondation_" + str(f.id())))
        ws.cell(row=row, column=2, value=f.is_a())
        ws.cell(row=row, column=3, value=round(volume, 3))
        total_vol_fond += volume
        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=3, value=round(total_vol_fond, 2)).font = Font(bold=True)
    ajuster_largeurs(ws, [25, 20, 14])

    print("  Fondations   : %d elements, %.2f m3" % (len(fondations), total_vol_fond))

    # ===== MENUISERIES =====
    ws = wb.create_sheet("Menuiseries")
    style_header(ws, ["Type", "Nom", "Etage", "Largeur (m)", "Hauteur (m)", "Surface (m²)"])

    row = 2
    nb_portes = 0
    nb_fenetres = 0

    for p in model.by_type("IfcDoor"):
        largeur = getattr(p, 'OverallWidth', None) or 0
        hauteur = getattr(p, 'OverallHeight', None) or 0
        try:
            largeur = float(largeur)
            hauteur = float(hauteur)
        except Exception:
            largeur, hauteur = 0, 0

        ws.cell(row=row, column=1, value="Porte")
        ws.cell(row=row, column=2, value=p.Name or ("Porte_" + str(p.id())))
        ws.cell(row=row, column=3, value=get_etage(p))
        ws.cell(row=row, column=4, value=round(largeur, 3))
        ws.cell(row=row, column=5, value=round(hauteur, 3))
        ws.cell(row=row, column=6, value=round(largeur * hauteur, 3))
        nb_portes += 1
        row += 1

    for f in model.by_type("IfcWindow"):
        largeur = getattr(f, 'OverallWidth', None) or 0
        hauteur = getattr(f, 'OverallHeight', None) or 0
        try:
            largeur = float(largeur)
            hauteur = float(hauteur)
        except Exception:
            largeur, hauteur = 0, 0

        ws.cell(row=row, column=1, value="Fenetre")
        ws.cell(row=row, column=2, value=f.Name or ("Fenetre_" + str(f.id())))
        ws.cell(row=row, column=3, value=get_etage(f))
        ws.cell(row=row, column=4, value=round(largeur, 3))
        ws.cell(row=row, column=5, value=round(hauteur, 3))
        ws.cell(row=row, column=6, value=round(largeur * hauteur, 3))
        nb_fenetres += 1
        row += 1

    ajuster_largeurs(ws, [10, 20, 12, 12, 12, 12])

    print("  Portes       : %d" % nb_portes)
    print("  Fenetres     : %d" % nb_fenetres)

    # ===== SYNTHESE (en premiere position) =====
    ws = wb.create_sheet("Synthese", 0)
    ws['A1'] = "SYNTHESE QUANTITATIVE"
    ws['A1'].font = Font(bold=True, size=14, color="1E3A5F")
    ws.merge_cells('A1:D1')

    headers = ["Ouvrage", "Unite", "Quantite", "Observation"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    data = [
        ("Fondations (toutes)", "m³", round(total_vol_fond, 2), "Semelles + pieux"),
        ("Murs - volume beton", "m³", round(total_vol_mur, 2), str(len(murs)) + " murs"),
        ("Murs - surface", "m²", round(total_surf_mur, 2), ""),
        ("Dalles - volume", "m³", round(total_vol_dalle, 2), str(len(dalles)) + " dalles"),
        ("Dalles - surface", "m²", round(total_surf_dalle, 2), ""),
        ("Poteaux - volume", "m³", round(total_vol_pot, 2), str(len(poteaux)) + " poteaux"),
        ("Poutres - volume", "m³", round(total_vol_pou, 2), str(len(poutres)) + " poutres"),
        ("", "", "", ""),
        ("TOTAL BETON ARME", "m³",
         round(total_vol_fond + total_vol_mur + total_vol_dalle + total_vol_pot + total_vol_pou, 2),
         "Somme - verification visuelle requise"),
        ("", "", "", ""),
        ("Portes", "U", nb_portes, ""),
        ("Fenetres", "U", nb_fenetres, ""),
    ]

    for i, (ouvrage, unite, qte, obs) in enumerate(data, 4):
        ws.cell(row=i, column=1, value=ouvrage)
        ws.cell(row=i, column=2, value=unite)
        ws.cell(row=i, column=3, value=qte)
        ws.cell(row=i, column=4, value=obs)
        if "TOTAL" in str(ouvrage):
            for col in range(1, 5):
                ws.cell(row=i, column=col).font = Font(bold=True)
                ws.cell(row=i, column=col).fill = PatternFill("solid", fgColor="FEF3C7")

    ajuster_largeurs(ws, [30, 10, 15, 35])

    # Sauvegarder
    base = os.path.splitext(chemin_ifc)[0]
    output = base + "_quantites.xlsx"

    try:
        wb.save(output)
    except Exception as e:
        print("[ERREUR] Impossible d'ecrire le fichier Excel : " + str(e))
        sys.exit(1)

    print("")
    print("[OK] Fichier genere : " + output)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 2_extract_quantites.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    extraire_quantites(sys.argv[1])

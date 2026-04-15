#!/usr/bin/env python3
"""
5_comparer_ifc.py — Comparer deux versions d'une maquette IFC
Usage : python 5_comparer_ifc.py maquette_v1.ifc maquette_v2.ifc

Produit : fichier Excel avec :
  - Elements ajoutes (presents dans v2 mais pas v1)
  - Elements supprimes (presents dans v1 mais pas v2)
  - Differences de quantites par type
  - Rapport ecart (%)
"""
import sys
import os


def comparer_ifc(chemin_v1: str, chemin_v2: str):
    try:
        import ifcopenshell
        import ifcopenshell.util.element
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError as e:
        print(f"[ERREUR] Bibliotheque manquante : {e}")
        sys.exit(1)

    if not os.path.exists(chemin_v1):
        print(f"[ERREUR] Fichier V1 introuvable : {chemin_v1}")
        sys.exit(1)
    if not os.path.exists(chemin_v2):
        print(f"[ERREUR] Fichier V2 introuvable : {chemin_v2}")
        sys.exit(1)

    print(f"\nComparaison :")
    print(f"  V1 : {os.path.basename(chemin_v1)}")
    print(f"  V2 : {os.path.basename(chemin_v2)}")

    m1 = ifcopenshell.open(chemin_v1)
    m2 = ifcopenshell.open(chemin_v2)

    def get_qto(element, prop_name):
        psets = ifcopenshell.util.element.get_psets(element, qtos_only=True)
        for qto_name, props in psets.items():
            if prop_name in props:
                return props[prop_name]
        return 0

    # Types a comparer
    types_compares = [
        "IfcWall", "IfcSlab", "IfcBeam", "IfcColumn",
        "IfcFooting", "IfcDoor", "IfcWindow", "IfcRoof",
        "IfcStair", "IfcRailing", "IfcCovering",
    ]

    # Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparaison"

    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1E3A5F")
    title_font = Font(bold=True, size=14, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1E3A5F")
    green_fill = PatternFill("solid", fgColor="DCFCE7")
    red_fill = PatternFill("solid", fgColor="FEE2E2")
    yellow_fill = PatternFill("solid", fgColor="FEF3C7")

    # En-tete
    ws.merge_cells('A1:F1')
    ws['A1'] = "COMPARAISON DE MAQUETTES IFC"
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws['A2'] = f"V1 : {os.path.basename(chemin_v1)}"
    ws['A3'] = f"V2 : {os.path.basename(chemin_v2)}"

    # Headers
    headers = ["Type d'element", "Nombre V1", "Nombre V2", "Ecart", "Variation %", "Observation"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    row = 6
    print(f"\n{'Type':20s} {'V1':>6s} {'V2':>6s} {'Ecart':>8s}  Observation")
    print("-" * 70)

    for type_ifc in types_compares:
        nb1 = len(m1.by_type(type_ifc))
        nb2 = len(m2.by_type(type_ifc))
        ecart = nb2 - nb1
        variation = (ecart / nb1 * 100) if nb1 > 0 else (100 if nb2 > 0 else 0)

        if nb1 == 0 and nb2 == 0:
            continue

        if ecart > 0:
            obs = f"+{ecart} element(s) ajoute(s)"
            fill = green_fill
        elif ecart < 0:
            obs = f"{ecart} element(s) supprime(s)"
            fill = red_fill
        else:
            obs = "Identique"
            fill = None

        ws.cell(row=row, column=1, value=type_ifc)
        ws.cell(row=row, column=2, value=nb1)
        ws.cell(row=row, column=3, value=nb2)
        ws.cell(row=row, column=4, value=ecart)
        ws.cell(row=row, column=5, value=f"{variation:+.1f}%")
        ws.cell(row=row, column=6, value=obs)

        if fill:
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = fill

        print(f"{type_ifc:20s} {nb1:>6d} {nb2:>6d} {ecart:>+8d}  {obs}")
        row += 1

    # Comparaison quantites
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=1, value="COMPARAISON DES QUANTITES (volumes et surfaces)")
    cell.font = title_font
    cell.fill = title_fill
    cell.alignment = Alignment(horizontal='center')
    row += 1

    headers2 = ["Grandeur", "Unite", "V1", "V2", "Ecart", "Variation %"]
    for i, h in enumerate(headers2, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    row += 1

    def total_volume(model, type_ifc):
        return sum((get_qto(e, 'NetVolume') or get_qto(e, 'GrossVolume') or 0)
                   for e in model.by_type(type_ifc))

    def total_surface(model, type_ifc, prop='NetArea'):
        alt = 'GrossArea' if prop == 'NetArea' else 'GrossSideArea'
        return sum((get_qto(e, prop) or get_qto(e, alt) or 0)
                   for e in model.by_type(type_ifc))

    grandeurs = [
        ("Volume murs", "m³", total_volume(m1, "IfcWall"), total_volume(m2, "IfcWall")),
        ("Volume dalles", "m³", total_volume(m1, "IfcSlab"), total_volume(m2, "IfcSlab")),
        ("Volume poteaux", "m³", total_volume(m1, "IfcColumn"), total_volume(m2, "IfcColumn")),
        ("Volume poutres", "m³", total_volume(m1, "IfcBeam"), total_volume(m2, "IfcBeam")),
        ("Volume fondations", "m³",
            total_volume(m1, "IfcFooting") + total_volume(m1, "IfcPile"),
            total_volume(m2, "IfcFooting") + total_volume(m2, "IfcPile")),
        ("Surface dalles", "m²",
            total_surface(m1, "IfcSlab", 'NetArea'),
            total_surface(m2, "IfcSlab", 'NetArea')),
        ("Surface murs", "m²",
            total_surface(m1, "IfcWall", 'NetSideArea'),
            total_surface(m2, "IfcWall", 'NetSideArea')),
    ]

    for nom, unite, v1, v2 in grandeurs:
        if v1 == 0 and v2 == 0:
            continue
        ecart = v2 - v1
        variation = (ecart / v1 * 100) if v1 > 0 else 100

        ws.cell(row=row, column=1, value=nom)
        ws.cell(row=row, column=2, value=unite)
        ws.cell(row=row, column=3, value=round(v1, 2))
        ws.cell(row=row, column=4, value=round(v2, 2))
        ws.cell(row=row, column=5, value=round(ecart, 2))
        ws.cell(row=row, column=6, value=f"{variation:+.1f}%")

        if abs(variation) > 10:
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = yellow_fill
        row += 1

    # Largeurs
    widths = [25, 12, 12, 12, 12, 30]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + col)].width = width

    # Sauvegarder
    base = os.path.splitext(chemin_v1)[0]
    output = base + "_vs_v2_comparaison.xlsx"
    wb.save(output)

    print(f"\n[OK] Comparaison generee : {output}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage : python 5_comparer_ifc.py <maquette_v1.ifc> <maquette_v2.ifc>")
        sys.exit(1)
    comparer_ifc(sys.argv[1], sys.argv[2])

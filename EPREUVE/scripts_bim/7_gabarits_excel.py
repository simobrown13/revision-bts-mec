#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
7_gabarits_excel.py - Genere des gabarits Excel modernes pour l'examen E6-A BIM
Usage : python 7_gabarits_excel.py [dossier_sortie]

Genere 5 gabarits pret-a-remplir :
  1. Dashboard projet (KPI globaux)
  2. DPGF professionnel (12 lots + totaux auto)
  3. Planning chantier (Gantt simplifie)
  4. Bilan carbone RE2020 (jauges + seuils)
  5. Controle qualite maquette BIM (checklist + score)
"""
import sys
import os

from _utils import setup_encoding

setup_encoding()

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule
    from openpyxl.chart import BarChart, PieChart, DoughnutChart, LineChart, Reference
    from openpyxl.chart.label import DataLabelList
except ImportError:
    print("[ERREUR] openpyxl manquant. Installez avec : python -m pip install openpyxl")
    sys.exit(1)


# ============================================================
# PALETTE COULEURS MODERNE
# ============================================================
NAVY = "1E3A5F"
NAVY_DARK = "0F1E33"
ORANGE = "F58220"
TURQUOISE = "4DC7C7"
GREEN = "16A34A"
RED = "DC2626"
YELLOW = "F59E0B"
GREY_LIGHT = "F1F5F9"
GREY = "94A3B8"
WHITE = "FFFFFF"


# ============================================================
# HELPERS STYLES
# ============================================================
def thin_border():
    s = Side(style='thin', color="CBD5E1")
    return Border(left=s, right=s, top=s, bottom=s)


def style_titre(ws, cell_range, texte, bg=NAVY, fg=WHITE, taille=16, hauteur=32):
    ws.merge_cells(cell_range)
    first = cell_range.split(':')[0]
    c = ws[first]
    c.value = texte
    c.font = Font(bold=True, size=taille, color=fg, name="Calibri")
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[c.row].height = hauteur


def style_sous_titre(ws, row, cell_range, texte, bg=ORANGE, fg=WHITE):
    ws.merge_cells(cell_range)
    first = cell_range.split(':')[0]
    c = ws[first]
    c.value = texte
    c.font = Font(bold=True, size=12, color=fg)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row].height = 22


def style_header(ws, row, headers, bg=NAVY, fg=WHITE):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(bold=True, color=fg, size=11)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
    ws.row_dimensions[row].height = 28


def kpi_card(ws, row, col, label, valeur, unite, couleur):
    """Cree une carte KPI moderne sur 2 lignes x 1 colonne."""
    # Ligne valeur
    c1 = ws.cell(row=row, column=col, value=valeur)
    c1.font = Font(bold=True, size=22, color=couleur)
    c1.alignment = Alignment(horizontal='center', vertical='center')
    c1.fill = PatternFill("solid", fgColor=GREY_LIGHT)
    ws.row_dimensions[row].height = 38

    # Ligne label
    c2 = ws.cell(row=row + 1, column=col, value=label + " (" + unite + ")")
    c2.font = Font(bold=True, size=10, color=NAVY)
    c2.alignment = Alignment(horizontal='center', vertical='center')
    c2.fill = PatternFill("solid", fgColor=GREY_LIGHT)
    ws.row_dimensions[row + 1].height = 22


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================
# GABARIT 1 : DASHBOARD PROJET
# ============================================================
def gabarit_dashboard(wb):
    ws = wb.create_sheet("1. Dashboard")

    style_titre(ws, 'A1:H1', "DASHBOARD PROJET BIM", bg=NAVY, taille=18, hauteur=40)

    ws['A2'] = "Projet :"
    ws['A2'].font = Font(bold=True, color=NAVY)
    ws['B2'] = "[A remplir]"
    ws['E2'] = "Phase :"
    ws['E2'].font = Font(bold=True, color=NAVY)
    ws['F2'] = "DCE"
    ws['A3'] = "Date :"
    ws['A3'].font = Font(bold=True, color=NAVY)
    ws['B3'] = "[JJ/MM/AAAA]"
    ws['E3'] = "Maitre d'oeuvre :"
    ws['E3'].font = Font(bold=True, color=NAVY)
    ws['F3'] = "[A remplir]"

    # Ligne KPI
    style_sous_titre(ws, 5, 'A5:H5', "INDICATEURS CLES DU PROJET", bg=ORANGE)

    kpi_card(ws, 7, 1, "Surface SHON", 0, "m²", NAVY)
    kpi_card(ws, 7, 2, "Volume beton", 0, "m³", ORANGE)
    kpi_card(ws, 7, 3, "Nb elements", 0, "u", TURQUOISE)
    kpi_card(ws, 7, 4, "Cout prev.", 0, "k€", GREEN)
    kpi_card(ws, 7, 5, "IC carbone", 0, "kgCO2/m²", RED)
    kpi_card(ws, 7, 6, "Delai", 0, "mois", NAVY)
    kpi_card(ws, 7, 7, "Nb lots", 12, "", ORANGE)
    kpi_card(ws, 7, 8, "Conformite", 0, "%", GREEN)

    # Tableau repartition lots
    style_sous_titre(ws, 11, 'A11:H11', "REPARTITION PAR LOT (€ HT)", bg=TURQUOISE)
    style_header(ws, 13, ["Code", "Lot", "Montant HT (€)", "% Total",
                           "Avancement %", "", "", ""])

    lots = [
        ("01", "Terrassement"), ("02", "Gros oeuvre"),
        ("03", "Charpente / Couverture"), ("04", "Menuiseries ext."),
        ("05", "Cloisons / Doublages"), ("06", "Revetements sols"),
        ("07", "Menuiseries int."), ("08", "Peinture"),
        ("09", "Plomberie"), ("10", "Electricite"),
        ("11", "CVC"), ("12", "VRD"),
    ]

    for i, (code, nom) in enumerate(lots):
        r = 14 + i
        ws.cell(row=r, column=1, value=code).alignment = Alignment(horizontal='center')
        ws.cell(row=r, column=2, value=nom)
        ws.cell(row=r, column=3, value=0).number_format = '#,##0 €'
        ws.cell(row=r, column=4, value="=IF(SUM($C$14:$C$25)=0,0,C" + str(r) + "/SUM($C$14:$C$25))")
        ws.cell(row=r, column=4).number_format = '0.0%'
        ws.cell(row=r, column=5, value=0).number_format = '0%'
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = thin_border()

    # Total
    r = 26
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True, size=12, color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=r, column=3, value="=SUM(C14:C25)").font = Font(bold=True, size=12)
    ws.cell(row=r, column=3).number_format = '#,##0 €'
    ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=YELLOW)
    ws.row_dimensions[r].height = 28

    # Data bar avancement
    ws.conditional_formatting.add("E14:E25",
        DataBarRule(start_type='num', start_value=0,
                    end_type='num', end_value=100, color=TURQUOISE))

    # Color scale % lot
    ws.conditional_formatting.add("D14:D25",
        ColorScaleRule(start_type='min', start_color="FFFFFF",
                       end_type='max', end_color=ORANGE))

    # Graphique donut lots
    chart = DoughnutChart()
    chart.title = "Repartition budget par lot"
    labels = Reference(ws, min_col=2, min_row=14, max_row=25)
    data = Reference(ws, min_col=3, min_row=13, max_row=25)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.height = 10
    chart.width = 15
    chart.dataLabels = DataLabelList(showPercent=True)
    ws.add_chart(chart, "J5")

    set_widths(ws, [10, 28, 18, 12, 15, 5, 5, 5])


# ============================================================
# GABARIT 2 : DPGF PROFESSIONNEL
# ============================================================
def gabarit_dpgf(wb):
    ws = wb.create_sheet("2. DPGF")

    style_titre(ws, 'A1:F1',
                "DECOMPOSITION DU PRIX GLOBAL ET FORFAITAIRE (DPGF)",
                bg=NAVY, taille=15)

    ws['A2'] = "Projet :"
    ws['A2'].font = Font(bold=True, color=NAVY)
    ws['B2'] = "[A remplir]"
    ws['A3'] = "Phase :"
    ws['A3'].font = Font(bold=True, color=NAVY)
    ws['B3'] = "DCE"

    style_header(ws, 5, ["Code", "Designation", "Unite", "Qte",
                          "PU HT (€)", "Total HT (€)"])

    lots_data = [
        ("LOT 01 - TERRASSEMENT", "8B5CF6", [
            ("01.01", "Decapage terre vegetale", "m²"),
            ("01.02", "Fouilles en pleine masse", "m³"),
            ("01.03", "Remblai compacte", "m³"),
        ]),
        ("LOT 02 - GROS OEUVRE", "3B82F6", [
            ("02.01", "Fondations semelles BA", "m³"),
            ("02.02", "Voiles BA", "m³"),
            ("02.03", "Dalles BA", "m³"),
            ("02.04", "Poteaux BA", "m³"),
            ("02.05", "Poutres BA", "m³"),
            ("02.06", "Maconnerie agglos", "m²"),
        ]),
        ("LOT 03 - CHARPENTE / COUVERTURE", "F59E0B", [
            ("03.01", "Charpente", "m²"),
            ("03.02", "Couverture", "m²"),
            ("03.03", "Zinguerie", "ml"),
        ]),
        ("LOT 04 - MENUISERIES EXTERIEURES", "22C55E", [
            ("04.01", "Fenetres", "U"),
            ("04.02", "Portes exterieures", "U"),
            ("04.03", "Occultations", "U"),
        ]),
        ("LOT 05 - CLOISONS / DOUBLAGES", "8B5CF6", [
            ("05.01", "Cloisons placo", "m²"),
            ("05.02", "Doublages isolants", "m²"),
            ("05.03", "Faux plafonds", "m²"),
        ]),
        ("LOT 06 - REVETEMENTS SOLS", "14B8A6", [
            ("06.01", "Chape / ragreage", "m²"),
            ("06.02", "Carrelage", "m²"),
            ("06.03", "Parquet / stratifie", "m²"),
        ]),
        ("LOT 07 - MENUISERIES INTERIEURES", "F59E0B", [
            ("07.01", "Portes interieures", "U"),
            ("07.02", "Placards", "ml"),
        ]),
        ("LOT 08 - PEINTURE", "EF4444", [
            ("08.01", "Peinture murs/plafonds", "m²"),
            ("08.02", "Peinture boiseries", "m²"),
        ]),
        ("LOT 09 - PLOMBERIE", "3B82F6", [
            ("09.01", "Alimentation EF/ECS", "Ft"),
            ("09.02", "Evacuations", "Ft"),
            ("09.03", "Appareils sanitaires", "Ens"),
        ]),
        ("LOT 10 - ELECTRICITE", "F59E0B", [
            ("10.01", "Tableau electrique", "U"),
            ("10.02", "Points lumineux", "U"),
            ("10.03", "Prises", "U"),
        ]),
        ("LOT 11 - CVC", "EF4444", [
            ("11.01", "Chauffage", "Ens"),
            ("11.02", "VMC", "Ens"),
        ]),
        ("LOT 12 - VRD", "22C55E", [
            ("12.01", "Voiries", "m²"),
            ("12.02", "Reseaux exterieurs", "Ft"),
        ]),
    ]

    row = 6
    first_data_row = None
    for nom_lot, couleur, items in lots_data:
        style_sous_titre(ws, row, "A" + str(row) + ":F" + str(row),
                         nom_lot, bg=couleur)
        row += 1
        for code, designation, unite in items:
            if first_data_row is None:
                first_data_row = row
            ws.cell(row=row, column=1, value=code).alignment = Alignment(horizontal='center')
            ws.cell(row=row, column=2, value=designation)
            ws.cell(row=row, column=3, value=unite).alignment = Alignment(horizontal='center')
            ws.cell(row=row, column=4, value=0)
            ws.cell(row=row, column=5, value=0).number_format = '#,##0.00 €'
            ws.cell(row=row, column=6,
                    value="=D" + str(row) + "*E" + str(row)).number_format = '#,##0.00 €'
            for col in range(1, 7):
                ws.cell(row=row, column=col).border = thin_border()
            row += 1

    # Totaux HT / TVA / TTC
    last_data_row = row - 1
    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="TOTAL HT").font = Font(bold=True, size=13, color=WHITE)
    ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', indent=2, vertical='center')
    total_ht = ws.cell(row=row, column=6,
                       value="=SUM(F" + str(first_data_row) + ":F" + str(last_data_row) + ")")
    total_ht.font = Font(bold=True, size=13)
    total_ht.number_format = '#,##0.00 €'
    total_ht.fill = PatternFill("solid", fgColor=YELLOW)
    ws.row_dimensions[row].height = 28
    ht_row = row
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="TVA 20%").font = Font(bold=True, size=12)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', indent=2)
    ws.cell(row=row, column=6, value="=F" + str(ht_row) + "*0.2").number_format = '#,##0.00 €'
    ws.cell(row=row, column=6).font = Font(bold=True, size=12)
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="TOTAL TTC").font = Font(bold=True, size=14, color=WHITE)
    ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=ORANGE)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', indent=2, vertical='center')
    ws.cell(row=row, column=6, value="=F" + str(ht_row) + "*1.2").number_format = '#,##0.00 €'
    ws.cell(row=row, column=6).font = Font(bold=True, size=14)
    ws.cell(row=row, column=6).fill = PatternFill("solid", fgColor=YELLOW)
    ws.row_dimensions[row].height = 32

    set_widths(ws, [10, 50, 10, 12, 15, 18])


# ============================================================
# GABARIT 3 : PLANNING CHANTIER
# ============================================================
def gabarit_planning(wb):
    ws = wb.create_sheet("3. Planning")

    style_titre(ws, 'A1:P1', "PLANNING CHANTIER (Gantt simplifie)",
                bg=NAVY, taille=15)

    ws['A2'] = "Debut :"
    ws['A2'].font = Font(bold=True, color=NAVY)
    ws['B2'] = "[JJ/MM/AAAA]"
    ws['E2'] = "Duree totale :"
    ws['E2'].font = Font(bold=True, color=NAVY)
    ws['F2'] = 12
    ws['G2'] = "mois"

    # En-tete tableau : tache + 12 mois
    headers = ["Tache", "Debut (S)", "Duree (S)"]
    for i in range(1, 13):
        headers.append("M" + str(i))
    style_header(ws, 4, headers)

    taches = [
        "Installation chantier",
        "Terrassement",
        "Fondations",
        "Gros oeuvre - RdC",
        "Gros oeuvre - Etages",
        "Charpente / Couverture",
        "Menuiseries ext.",
        "Cloisons / Doublages",
        "Plomberie / Electricite / CVC",
        "Revetements sols",
        "Menuiseries int. / Peinture",
        "VRD / Exterieurs",
        "Reception / OPR",
    ]

    fill_orange = PatternFill("solid", fgColor=ORANGE)

    for i, t in enumerate(taches):
        r = 5 + i
        ws.cell(row=r, column=1, value=t)
        ws.cell(row=r, column=2, value=0)  # Debut en semaines
        ws.cell(row=r, column=3, value=0)  # Duree en semaines
        for col in range(1, 16):
            ws.cell(row=r, column=col).border = thin_border()
            if col >= 4:
                # Formule : colore la cellule si le mois est dans la plage [debut, debut+duree]
                # Mois n = semaines [4*(n-1)+1 .. 4*n]
                mois = col - 3
                debut_col = "$B" + str(r)
                duree_col = "$C" + str(r)
                # Rempli si (debut < 4*mois) ET (debut + duree > 4*(mois-1))
                formule = ('=IF(AND(' + debut_col + '<' + str(4 * mois) +
                           ',' + debut_col + '+' + duree_col + '>' + str(4 * (mois - 1)) +
                           '),"X","")')
                c = ws.cell(row=r, column=col, value=formule)
                c.alignment = Alignment(horizontal='center')
                c.font = Font(color=ORANGE, bold=True)

        ws.row_dimensions[r].height = 22

    # Format conditionnel : colore si "X"
    ws.conditional_formatting.add("D5:O" + str(4 + len(taches)),
        CellIsRule(operator='equal', formula=['"X"'], fill=fill_orange,
                   font=Font(color=ORANGE, bold=True)))

    set_widths(ws, [35, 10, 10] + [6] * 12 + [5])


# ============================================================
# GABARIT 4 : BILAN CARBONE RE2020
# ============================================================
def gabarit_bilan_carbone(wb):
    ws = wb.create_sheet("4. Bilan carbone")

    style_titre(ws, 'A1:F1', "BILAN CARBONE RE2020 - IC Construction",
                bg=GREEN, taille=15)

    ws['A2'] = "Projet :"
    ws['A2'].font = Font(bold=True, color=NAVY)
    ws['B2'] = "[A remplir]"
    ws['A3'] = "SHON :"
    ws['A3'].font = Font(bold=True, color=NAVY)
    ws['B3'] = 100
    ws['C3'] = "m²"

    # KPI cards
    style_sous_titre(ws, 5, 'A5:F5', "INDICATEURS CARBONE", bg=ORANGE)

    kpi_card(ws, 7, 1, "IC total", "=F50", "kgCO2eq", RED)
    kpi_card(ws, 7, 2, "IC / m²", "=F50/B3", "kgCO2eq/m²", ORANGE)
    kpi_card(ws, 7, 3, "Seuil 2025", 530, "kgCO2eq/m²", NAVY)
    kpi_card(ws, 7, 4, "Ecart seuil", "=F50/B3-530", "kgCO2eq/m²", NAVY)
    kpi_card(ws, 7, 5, "Conformite", '=IF(F50/B3<=530,"OK","DEPASSE")', "", GREEN)

    # Tableau lots carbone
    style_sous_titre(ws, 11, 'A11:F11', "DETAIL PAR LOT", bg=TURQUOISE)
    style_header(ws, 13, ["Lot", "Ouvrage", "Qte", "Unite",
                           "Facteur (kgCO2/U)", "kgCO2eq"])

    FDES = {
        "Beton": 250, "Acier": 2.0, "Bois charpente": -25,
        "Couverture": 18, "PVC menuiserie": 90, "Placo": 3,
        "Laine verre": 6, "PSE": 12,
    }

    lignes = [
        ("GROS OEUVRE", "Beton fondations", "m³", FDES["Beton"]),
        ("GROS OEUVRE", "Beton murs/voiles", "m³", FDES["Beton"]),
        ("GROS OEUVRE", "Beton dalles", "m³", FDES["Beton"]),
        ("GROS OEUVRE", "Beton poteaux/poutres", "m³", FDES["Beton"]),
        ("GROS OEUVRE", "Acier BA", "kg", FDES["Acier"]),
        ("CHARPENTE", "Charpente bois", "m²", FDES["Bois charpente"]),
        ("CHARPENTE", "Couverture", "m²", FDES["Couverture"]),
        ("MENUISERIES", "Fenetres PVC", "m²", FDES["PVC menuiserie"]),
        ("CLOISONS", "Cloisons placo", "m²", FDES["Placo"]),
        ("ISOLATION", "Laine de verre", "m²", FDES["Laine verre"]),
        ("ISOLATION", "PSE facade", "m²", FDES["PSE"]),
    ]

    row = 14
    for lot, ouvrage, unite, facteur in lignes:
        ws.cell(row=row, column=1, value=lot)
        ws.cell(row=row, column=2, value=ouvrage)
        ws.cell(row=row, column=3, value=0)  # Quantite a remplir
        ws.cell(row=row, column=4, value=unite).alignment = Alignment(horizontal='center')
        ws.cell(row=row, column=5, value=facteur)
        ws.cell(row=row, column=6, value="=C" + str(row) + "*E" + str(row))
        ws.cell(row=row, column=6).number_format = '#,##0'
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = thin_border()
        row += 1

    # TOTAL (ligne 50 pour les KPI)
    while row < 50:
        row += 1

    ws.merge_cells('A50:E50')
    ws['A50'] = "TOTAL IC CONSTRUCTION (kgCO2eq)"
    ws['A50'].font = Font(bold=True, size=13, color=WHITE)
    ws['A50'].fill = PatternFill("solid", fgColor=RED)
    ws['A50'].alignment = Alignment(horizontal='right', indent=2, vertical='center')
    ws['F50'] = "=SUM(F14:F24)"
    ws['F50'].font = Font(bold=True, size=13)
    ws['F50'].number_format = '#,##0'
    ws['F50'].fill = PatternFill("solid", fgColor=YELLOW)
    ws.row_dimensions[50].height = 30

    # Tableau seuils RE2020
    style_sous_titre(ws, 52, 'A52:F52', "SEUILS RE2020 (logement)", bg=NAVY)
    style_header(ws, 53, ["Periode", "Seuil (kgCO2/m²)", "Projet (kgCO2/m²)",
                           "Ecart", "Statut", ""])

    seuils = [
        ("2022-2024", 640), ("2025-2027", 530),
        ("2028-2030", 475), ("2031+", 415),
    ]
    for i, (periode, seuil) in enumerate(seuils):
        r = 54 + i
        ws.cell(row=r, column=1, value=periode)
        ws.cell(row=r, column=2, value=seuil)
        ws.cell(row=r, column=3, value="=$F$50/$B$3").number_format = '0'
        ws.cell(row=r, column=4, value="=C" + str(r) + "-B" + str(r)).number_format = '+0;-0;0'
        ws.cell(row=r, column=5,
                value='=IF(C' + str(r) + '<=B' + str(r) + ',"Respecte","Depasse")')
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = thin_border()
            ws.cell(row=r, column=col).alignment = Alignment(horizontal='center')

    # Format conditionnel statut
    ws.conditional_formatting.add("E54:E57",
        CellIsRule(operator='equal', formula=['"Respecte"'],
                   fill=PatternFill("solid", fgColor="DCFCE7"),
                   font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add("E54:E57",
        CellIsRule(operator='equal', formula=['"Depasse"'],
                   fill=PatternFill("solid", fgColor="FEE2E2"),
                   font=Font(color=RED, bold=True)))

    # Graphique barres seuils
    chart = BarChart()
    chart.type = "bar"
    chart.style = 11
    chart.title = "Comparaison seuils RE2020"
    chart.y_axis.title = 'kgCO2eq/m²'
    data = Reference(ws, min_col=2, min_row=53, max_col=3, max_row=57)
    cats = Reference(ws, min_col=1, min_row=54, max_row=57)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 16
    ws.add_chart(chart, "H11")

    set_widths(ws, [18, 28, 12, 10, 18, 15])


# ============================================================
# GABARIT 5 : CONTROLE QUALITE MAQUETTE BIM
# ============================================================
def gabarit_controle_qualite(wb):
    ws = wb.create_sheet("5. Controle BIM")

    style_titre(ws, 'A1:F1', "CONTROLE QUALITE MAQUETTE BIM",
                bg=TURQUOISE, fg=WHITE, taille=15)

    ws['A2'] = "Maquette :"
    ws['A2'].font = Font(bold=True, color=NAVY)
    ws['B2'] = "[fichier.ifc]"
    ws['A3'] = "Controleur :"
    ws['A3'].font = Font(bold=True, color=NAVY)
    ws['B3'] = "[nom]"

    # KPI score global
    style_sous_titre(ws, 5, 'A5:F5', "SCORE QUALITE", bg=ORANGE)

    kpi_card(ws, 7, 1, "Score total", "=COUNTIF(D13:D50,\"OK\")", "/ 30", GREEN)
    kpi_card(ws, 7, 2, "Non conformes", "=COUNTIF(D13:D50,\"KO\")", "", RED)
    kpi_card(ws, 7, 3, "A verifier", "=COUNTIF(D13:D50,\"?\")", "", YELLOW)
    kpi_card(ws, 7, 4, "Conformite",
             "=IFERROR(COUNTIF(D13:D50,\"OK\")/(COUNTA(D13:D50)),0)", "%", NAVY)

    # Checklist
    style_sous_titre(ws, 11, 'A11:F11', "CHECKLIST CONTROLE", bg=NAVY)
    style_header(ws, 12, ["Categorie", "Point de controle", "Critere attendu",
                           "Statut", "Commentaire", ""])

    controles = [
        ("Structure IFC", "Fichier IFC valide", "Ouverture sans erreur"),
        ("Structure IFC", "Version IFC", "IFC2x3 ou IFC4"),
        ("Structure IFC", "Unites declarees", "Metres, m², m³"),
        ("Structure IFC", "Projet/Site/Batiment/Etages", "Hierarchie complete"),
        ("Geometrie", "Murs presents", "IfcWall non vide"),
        ("Geometrie", "Dalles presentes", "IfcSlab non vide"),
        ("Geometrie", "Poteaux/Poutres", "IfcColumn/IfcBeam"),
        ("Geometrie", "Fondations", "IfcFooting/IfcPile"),
        ("Geometrie", "Menuiseries", "IfcDoor/IfcWindow"),
        ("Geometrie", "Toiture", "IfcRoof"),
        ("Quantites", "Volumes renseignes", "NetVolume > 0"),
        ("Quantites", "Surfaces renseignees", "NetArea > 0"),
        ("Quantites", "Longueurs renseignees", "Length > 0"),
        ("Psets", "Pset_WallCommon", "LoadBearing, IsExternal"),
        ("Psets", "Pset_SlabCommon", "LoadBearing, PitchAngle"),
        ("Psets", "Pset_DoorCommon", "FireRating, IsExternal"),
        ("Psets", "Pset_WindowCommon", "GlazingAreaFraction"),
        ("Materiaux", "Materiau beton defini", "IfcMaterial assigne"),
        ("Materiaux", "Materiau isolant defini", "IfcMaterial assigne"),
        ("Materiaux", "Materiau menuiserie", "IfcMaterial assigne"),
        ("Classification", "Code Uniformat/Omniclass", "Optionnel"),
        ("Classification", "Identifiants uniques", "GlobalId present"),
        ("Nommage", "Noms coherents", "Mur_RDC_01, etc."),
        ("Nommage", "Types definis", "IfcWallType, etc."),
        ("Etages", "Noms etages explicites", "RDC, R+1, etc."),
        ("Etages", "Elements affectes etage", "ContainedInStructure"),
        ("Coherence", "Pas de doublons", "GUID uniques"),
        ("Coherence", "Elements orphelins", "Tous relies au batiment"),
        ("Livrables", "Nomenclature quantites", "Excel/BCF"),
        ("Livrables", "Rapport de controle", "PDF/Excel"),
    ]

    for i, (cat, point, critere) in enumerate(controles):
        r = 13 + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=point)
        ws.cell(row=r, column=3, value=critere)
        ws.cell(row=r, column=4, value="").alignment = Alignment(horizontal='center')
        ws.cell(row=r, column=5, value="")
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = thin_border()
        ws.row_dimensions[r].height = 20

    # Format conditionnel statut
    statut_range = "D13:D" + str(12 + len(controles))
    ws.conditional_formatting.add(statut_range,
        CellIsRule(operator='equal', formula=['"OK"'],
                   fill=PatternFill("solid", fgColor="DCFCE7"),
                   font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(statut_range,
        CellIsRule(operator='equal', formula=['"KO"'],
                   fill=PatternFill("solid", fgColor="FEE2E2"),
                   font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(statut_range,
        CellIsRule(operator='equal', formula=['"?"'],
                   fill=PatternFill("solid", fgColor="FEF3C7"),
                   font=Font(color=YELLOW, bold=True)))

    # Validation donnees : OK / KO / ?
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"OK,KO,?"', allow_blank=True)
    dv.add(statut_range)
    ws.add_data_validation(dv)

    set_widths(ws, [18, 32, 32, 12, 35, 5])


# ============================================================
# PAGE GARDE
# ============================================================
def page_garde(wb):
    ws = wb.active
    ws.title = "Sommaire"

    style_titre(ws, 'A1:D1', "GABARITS BIM - BTS MEC 2026",
                bg=NAVY, taille=22, hauteur=60)

    ws['A3'] = "Candidat :"
    ws['A3'].font = Font(bold=True, size=12, color=NAVY)
    ws['B3'] = "BAHAFID Mohamed"
    ws['A4'] = "Epreuve :"
    ws['A4'].font = Font(bold=True, size=12, color=NAVY)
    ws['B4'] = "E6-A Projet numerique BIM"
    ws['A5'] = "Date :"
    ws['A5'].font = Font(bold=True, size=12, color=NAVY)
    ws['B5'] = "[a remplir]"

    style_sous_titre(ws, 7, 'A7:D7', "SOMMAIRE DES GABARITS", bg=ORANGE)

    gabarits = [
        ("1", "Dashboard projet", "KPI globaux, repartition lots, graphique donut"),
        ("2", "DPGF professionnel", "12 lots, totaux HT/TVA/TTC automatiques"),
        ("3", "Planning chantier", "Gantt 12 mois avec remplissage auto"),
        ("4", "Bilan carbone RE2020", "IC construction, seuils, graphique"),
        ("5", "Controle qualite BIM", "Checklist 30 points, score conformite"),
    ]

    style_header(ws, 9, ["N°", "Gabarit", "Contenu", ""])

    for i, (num, nom, contenu) in enumerate(gabarits):
        r = 10 + i
        ws.cell(row=r, column=1, value=num).alignment = Alignment(horizontal='center')
        ws.cell(row=r, column=1).font = Font(bold=True, size=14, color=ORANGE)
        ws.cell(row=r, column=2, value=nom).font = Font(bold=True, color=NAVY)
        ws.cell(row=r, column=3, value=contenu)
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = thin_border()
        ws.row_dimensions[r].height = 26

    ws['A18'] = "Mode d'emploi :"
    ws['A18'].font = Font(bold=True, size=12, color=NAVY)
    ws['A19'] = "1. Completer les champs [A remplir] en en-tete de chaque onglet"
    ws['A20'] = "2. Saisir les quantites extraites de la maquette eveBIM"
    ws['A21'] = "3. Les totaux, pourcentages et indicateurs se calculent automatiquement"
    ws['A22'] = "4. Les graphiques et jauges se mettent a jour en temps reel"
    ws['A23'] = "5. Imprimer en A4 paysage pour les annexes"

    ws['A25'] = "Astuce examen :"
    ws['A25'].font = Font(bold=True, size=12, color=ORANGE)
    ws['A26'] = "Utiliser les scripts 2_extract_quantites.py et 3_generer_dpgf.py pour"
    ws['A27'] = "remplir automatiquement ces gabarits a partir du fichier IFC."

    set_widths(ws, [8, 30, 55, 5])


# ============================================================
# MAIN
# ============================================================
def generer_gabarits(dossier_sortie):
    if not os.path.isdir(dossier_sortie):
        os.makedirs(dossier_sortie, exist_ok=True)

    wb = Workbook()
    page_garde(wb)
    gabarit_dashboard(wb)
    gabarit_dpgf(wb)
    gabarit_planning(wb)
    gabarit_bilan_carbone(wb)
    gabarit_controle_qualite(wb)

    output = os.path.join(dossier_sortie, "GABARITS_BIM_Examen.xlsx")

    try:
        wb.save(output)
    except Exception as e:
        print("[ERREUR] Impossible d'ecrire : " + str(e))
        sys.exit(1)

    print("")
    print("[OK] Gabarits generes : " + output)
    print("")
    print("Contenu :")
    print("  Sommaire")
    print("  1. Dashboard projet (KPI + donut chart)")
    print("  2. DPGF (12 lots + totaux auto)")
    print("  3. Planning chantier (Gantt 12 mois)")
    print("  4. Bilan carbone RE2020 (jauges + seuils)")
    print("  5. Controle qualite BIM (checklist 30 points)")
    print("")
    print("Pret a utiliser pour l'examen E6-A BIM.")


if __name__ == "__main__":
    dossier = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    generer_gabarits(dossier)

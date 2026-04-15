#!/usr/bin/env python3
"""
1_analyse_ifc.py — Analyser la structure d'une maquette IFC
Usage : python 1_analyse_ifc.py chemin/vers/maquette.ifc

Produit : rapport texte avec :
  - Informations generales (projet, site, batiment)
  - Nombre d'elements par type
  - Etages et espaces
  - Arborescence simplifiee
"""
import sys
import os
from collections import Counter


def analyser_ifc(chemin_ifc: str):
    try:
        import ifcopenshell
    except ImportError:
        print("[ERREUR] ifcopenshell non installe.")
        print("Lancez : pip install ifcopenshell")
        sys.exit(1)

    if not os.path.exists(chemin_ifc):
        print(f"[ERREUR] Fichier introuvable : {chemin_ifc}")
        sys.exit(1)

    print(f"\n{'=' * 60}")
    print(f"ANALYSE IFC : {os.path.basename(chemin_ifc)}")
    print(f"{'=' * 60}\n")

    # Charger la maquette
    model = ifcopenshell.open(chemin_ifc)

    # Informations generales
    projet = model.by_type("IfcProject")
    if projet:
        p = projet[0]
        print(f"Projet        : {p.Name or '(sans nom)'}")
        print(f"Description   : {p.Description or '(aucune)'}")
        print(f"Schema IFC    : {model.schema}")

    sites = model.by_type("IfcSite")
    if sites:
        print(f"\nSite          : {sites[0].Name or '(sans nom)'}")

    batiments = model.by_type("IfcBuilding")
    for b in batiments:
        print(f"Batiment      : {b.Name or '(sans nom)'}")

    etages = model.by_type("IfcBuildingStorey")
    print(f"\nETAGES ({len(etages)}) :")
    for e in etages:
        elev = getattr(e, 'Elevation', 'N/C')
        print(f"  - {e.Name or '(sans nom)':20s} altitude = {elev}")

    espaces = model.by_type("IfcSpace")
    print(f"\nESPACES/LOCAUX : {len(espaces)}")

    # Compter les elements par type
    print(f"\n{'=' * 60}")
    print("INVENTAIRE DES ELEMENTS")
    print(f"{'=' * 60}")

    types_a_compter = [
        ("IfcWall", "Murs (tous)"),
        ("IfcWallStandardCase", "  dont murs standards"),
        ("IfcSlab", "Dalles"),
        ("IfcBeam", "Poutres"),
        ("IfcColumn", "Poteaux"),
        ("IfcFooting", "Fondations"),
        ("IfcPile", "Pieux"),
        ("IfcDoor", "Portes"),
        ("IfcWindow", "Fenetres"),
        ("IfcRoof", "Toitures"),
        ("IfcStair", "Escaliers"),
        ("IfcRamp", "Rampes"),
        ("IfcCovering", "Revetements"),
        ("IfcCurtainWall", "Murs-rideaux"),
        ("IfcRailing", "Garde-corps"),
        ("IfcFurnishingElement", "Mobilier"),
        ("IfcSanitaryTerminal", "Sanitaires"),
    ]

    print(f"\n{'Type':35s} {'Nombre':>8s}")
    print("-" * 45)
    total = 0
    for type_ifc, label in types_a_compter:
        elements = model.by_type(type_ifc)
        nb = len(elements)
        if nb > 0:
            print(f"{label:35s} {nb:>8d}")
            if not label.startswith("  "):
                total += nb

    print("-" * 45)
    print(f"{'TOTAL (hors sous-categories)':35s} {total:>8d}")

    # Materiaux utilises
    materiaux = model.by_type("IfcMaterial")
    if materiaux:
        print(f"\n{'=' * 60}")
        print(f"MATERIAUX ({len(materiaux)})")
        print(f"{'=' * 60}")
        for m in sorted(materiaux, key=lambda x: x.Name or ''):
            if m.Name:
                print(f"  - {m.Name}")

    # Elements porteurs vs non porteurs
    print(f"\n{'=' * 60}")
    print("CLASSIFICATION STRUCTURELLE")
    print(f"{'=' * 60}")

    murs = model.by_type("IfcWall")
    porteurs = 0
    non_porteurs = 0
    exterieurs = 0
    interieurs = 0
    for mur in murs:
        psets = ifcopenshell.util.element.get_psets(mur) if hasattr(ifcopenshell, 'util') else {}
        common = psets.get('Pset_WallCommon', {})
        if common.get('LoadBearing'):
            porteurs += 1
        else:
            non_porteurs += 1
        if common.get('IsExternal'):
            exterieurs += 1
        else:
            interieurs += 1

    if murs:
        print(f"  Murs porteurs      : {porteurs}")
        print(f"  Murs non porteurs  : {non_porteurs}")
        print(f"  Murs exterieurs    : {exterieurs}")
        print(f"  Murs interieurs    : {interieurs}")

    print(f"\n{'=' * 60}")
    print("Analyse terminee")
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 1_analyse_ifc.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    analyser_ifc(sys.argv[1])

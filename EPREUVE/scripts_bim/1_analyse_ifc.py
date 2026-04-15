#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1_analyse_ifc.py - Analyser la structure d'une maquette IFC
Usage : python 1_analyse_ifc.py chemin/vers/maquette.ifc
"""
import sys
import os

# Force UTF-8 output on Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass


def get_psets_safe(element):
    """Recupere les Psets d'un element sans dependre de ifcopenshell.util."""
    psets = {}
    try:
        for rel in element.IsDefinedBy or []:
            if rel.is_a('IfcRelDefinesByProperties'):
                pset = rel.RelatingPropertyDefinition
                if pset.is_a('IfcPropertySet'):
                    props = {}
                    for prop in pset.HasProperties or []:
                        if prop.is_a('IfcPropertySingleValue') and prop.NominalValue:
                            props[prop.Name] = prop.NominalValue.wrappedValue
                    psets[pset.Name] = props
                elif pset.is_a('IfcElementQuantity'):
                    props = {}
                    for q in pset.Quantities or []:
                        try:
                            if q.is_a('IfcQuantityLength'):
                                props[q.Name] = q.LengthValue
                            elif q.is_a('IfcQuantityArea'):
                                props[q.Name] = q.AreaValue
                            elif q.is_a('IfcQuantityVolume'):
                                props[q.Name] = q.VolumeValue
                            elif q.is_a('IfcQuantityCount'):
                                props[q.Name] = q.CountValue
                            elif q.is_a('IfcQuantityWeight'):
                                props[q.Name] = q.WeightValue
                        except Exception:
                            pass
                    psets[pset.Name] = props
    except Exception:
        pass
    return psets


def get_container_safe(element):
    """Trouve le conteneur (etage) d'un element."""
    try:
        for rel in element.ContainedInStructure or []:
            return rel.RelatingStructure
    except Exception:
        pass
    return None


def analyser_ifc(chemin_ifc):
    try:
        import ifcopenshell
    except ImportError:
        print("[ERREUR] ifcopenshell non installe.")
        print("Lancez : pip install ifcopenshell")
        sys.exit(1)

    if not os.path.exists(chemin_ifc):
        print("[ERREUR] Fichier introuvable : " + chemin_ifc)
        sys.exit(1)

    print("")
    print("=" * 60)
    print("ANALYSE IFC : " + os.path.basename(chemin_ifc))
    print("=" * 60)
    print("")

    try:
        model = ifcopenshell.open(chemin_ifc)
    except Exception as e:
        print("[ERREUR] Impossible d'ouvrir le fichier IFC : " + str(e))
        sys.exit(1)

    # Informations generales
    projets = model.by_type("IfcProject")
    if projets:
        p = projets[0]
        print("Projet        : " + str(p.Name or "(sans nom)"))
        print("Description   : " + str(p.Description or "(aucune)"))
        print("Schema IFC    : " + model.schema)

    sites = model.by_type("IfcSite")
    if sites:
        print("Site          : " + str(sites[0].Name or "(sans nom)"))

    batiments = model.by_type("IfcBuilding")
    for b in batiments:
        print("Batiment      : " + str(b.Name or "(sans nom)"))

    etages = model.by_type("IfcBuildingStorey")
    print("")
    print("ETAGES (%d) :" % len(etages))
    for e in etages:
        elev = getattr(e, 'Elevation', None)
        elev_str = str(elev) if elev is not None else "N/C"
        print("  - " + str(e.Name or "(sans nom)").ljust(20) + " altitude = " + elev_str)

    espaces = model.by_type("IfcSpace")
    print("")
    print("ESPACES/LOCAUX : %d" % len(espaces))

    # Inventaire
    print("")
    print("=" * 60)
    print("INVENTAIRE DES ELEMENTS")
    print("=" * 60)

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

    print("")
    print("%-35s %8s" % ("Type", "Nombre"))
    print("-" * 45)
    total = 0
    for type_ifc, label in types_a_compter:
        try:
            elements = model.by_type(type_ifc)
            nb = len(elements)
            if nb > 0:
                print("%-35s %8d" % (label, nb))
                if not label.startswith("  "):
                    total += nb
        except Exception:
            pass

    print("-" * 45)
    print("%-35s %8d" % ("TOTAL (hors sous-categories)", total))

    # Materiaux
    try:
        materiaux = model.by_type("IfcMaterial")
        if materiaux:
            print("")
            print("=" * 60)
            print("MATERIAUX (%d)" % len(materiaux))
            print("=" * 60)
            noms = sorted([m.Name for m in materiaux if m.Name])
            for n in noms:
                print("  - " + n)
    except Exception:
        pass

    # Classification structurelle
    print("")
    print("=" * 60)
    print("CLASSIFICATION STRUCTURELLE")
    print("=" * 60)

    murs = model.by_type("IfcWall")
    porteurs = 0
    non_porteurs = 0
    exterieurs = 0
    interieurs = 0
    inconnus = 0

    for mur in murs:
        psets = get_psets_safe(mur)
        common = psets.get('Pset_WallCommon', {})
        lb = common.get('LoadBearing')
        ie = common.get('IsExternal')
        if lb is True:
            porteurs += 1
        elif lb is False:
            non_porteurs += 1
        else:
            inconnus += 1
        if ie is True:
            exterieurs += 1
        elif ie is False:
            interieurs += 1

    if murs:
        print("  Murs porteurs      : %d" % porteurs)
        print("  Murs non porteurs  : %d" % non_porteurs)
        print("  Murs classification inconnue : %d" % inconnus)
        print("  Murs exterieurs    : %d" % exterieurs)
        print("  Murs interieurs    : %d" % interieurs)

    print("")
    print("=" * 60)
    print("Analyse terminee")
    print("=" * 60)
    print("")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage : python 1_analyse_ifc.py <chemin_vers_maquette.ifc>")
        sys.exit(1)
    analyser_ifc(sys.argv[1])

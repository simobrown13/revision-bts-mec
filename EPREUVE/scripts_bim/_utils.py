#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
_utils.py - Fonctions utilitaires partagees par tous les scripts BIM
Compatible toutes versions de ifcopenshell (0.7.x et superieures)
"""
import sys


def setup_encoding():
    """Force UTF-8 output sur Windows."""
    if sys.platform == 'win32':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except Exception:
            pass


def check_dependencies():
    """Verifie que les dependances sont installees."""
    missing = []
    try:
        import ifcopenshell
    except ImportError:
        missing.append("ifcopenshell")
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    if missing:
        print("[ERREUR] Bibliotheques manquantes : " + ", ".join(missing))
        print("")
        print("Installation :")
        print("  python -m pip install " + " ".join(missing))
        print("")
        print("Ou sur Windows : double-cliquez sur INSTALL_Windows.bat")
        sys.exit(1)


def get_psets(element):
    """Recupere tous les Psets d'un element (sans ifcopenshell.util)."""
    psets = {}
    try:
        defined_by = element.IsDefinedBy or []
    except Exception:
        return psets

    for rel in defined_by:
        try:
            if not rel.is_a('IfcRelDefinesByProperties'):
                continue
            pset = rel.RelatingPropertyDefinition
            if pset.is_a('IfcPropertySet'):
                props = {}
                for prop in pset.HasProperties or []:
                    try:
                        if prop.is_a('IfcPropertySingleValue') and prop.NominalValue:
                            props[prop.Name] = prop.NominalValue.wrappedValue
                    except Exception:
                        pass
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


def get_qto(element, prop_name):
    """Extrait une quantite depuis les QuantitySets. Retourne 0 si non trouvee."""
    psets = get_psets(element)
    for pset_name, props in psets.items():
        # Chercher dans les QuantitySets (commencent souvent par Qto_)
        if prop_name in props:
            val = props[prop_name]
            try:
                return float(val) if val is not None else 0
            except Exception:
                return 0
    return 0


def get_pset_value(element, pset_name, prop_name):
    """Extrait une valeur d'un PropertySet specifique."""
    psets = get_psets(element)
    return psets.get(pset_name, {}).get(prop_name)


def get_etage(element):
    """Trouve le nom de l'etage d'un element."""
    try:
        for rel in element.ContainedInStructure or []:
            struct = rel.RelatingStructure
            if struct.is_a('IfcBuildingStorey'):
                return struct.Name or "N/C"
    except Exception:
        pass
    return "N/C"


def get_material(model, element):
    """Trouve le materiau principal d'un element."""
    try:
        for rel in model.by_type('IfcRelAssociatesMaterial'):
            if element in (rel.RelatedObjects or []):
                mat = rel.RelatingMaterial
                if hasattr(mat, 'Name') and mat.Name:
                    return mat.Name
                if hasattr(mat, 'ForLayerSet'):
                    # MaterialLayerSetUsage
                    for layer in mat.ForLayerSet.MaterialLayers or []:
                        if layer.Material and layer.Material.Name:
                            return layer.Material.Name
    except Exception:
        pass
    return ""


def safe_volume(element):
    """Retourne le volume d'un element (plusieurs tentatives)."""
    return (get_qto(element, 'NetVolume') or
            get_qto(element, 'GrossVolume') or
            get_qto(element, 'Volume') or 0)


def safe_area(element):
    """Retourne la surface d'un element."""
    return (get_qto(element, 'NetArea') or
            get_qto(element, 'GrossArea') or
            get_qto(element, 'NetSideArea') or
            get_qto(element, 'GrossSideArea') or
            get_qto(element, 'Area') or 0)


def safe_length(element):
    """Retourne la longueur d'un element."""
    return (get_qto(element, 'Length') or
            get_qto(element, 'NetLength') or
            get_qto(element, 'GrossLength') or 0)


def safe_height(element):
    """Retourne la hauteur d'un element."""
    return (get_qto(element, 'Height') or
            get_qto(element, 'NetHeight') or
            get_qto(element, 'GrossHeight') or 0)

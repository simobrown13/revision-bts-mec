# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════╗
║  CONTENU DU RAPPORT U62 – BAHAFID Mohamed                  ║
║  BTS Management Économique de la Construction – Session 2026 ║
╠══════════════════════════════════════════════════════════════╣
║  CE FICHIER CONTIENT TOUT LE TEXTE DU RAPPORT.              ║
║                                                              ║
║  Pour modifier le rapport :                                  ║
║  1. Changez le texte entre guillemets " " ci-dessous         ║
║  2. Relancez : python generer_pptx_30pages.py                ║
║                                                              ║
║  Les \n créent un saut de ligne.                             ║
╚══════════════════════════════════════════════════════════════╝
"""

# =============================================================================
# INFORMATIONS CANDIDAT (utilisées sur plusieurs pages)
# =============================================================================
CANDIDAT = {
    "nom": "BAHAFID Mohamed",
    "numero": "02537399911",
    "academie": "Lyon",
    "session": "2026",
    "structure": "Conseil Régional de Béni Mellal-Khénifra (Maroc)",
    "direction": "Agence d'Exécution des Projets",
    "poste": "Technicien Études et Suivi des Travaux",
    "experience": "8 ans dans le BTP (3 ans Maroc + 5 ans France)",
    "formation_bim": "Technicien Modeleur BIM – AFPA Colmar (8 mois)",
    "activite": "BIMCO – Projeteur BIM / Économiste de la construction",
    "siren": "999580053",
    "siret": "99958005300018",
    "ape": "7112B – Ingénierie, études techniques",
    "siege": "44 rue de la République, 42510 Bussières (Loire)",
    "site_web": "gestion.bimco-consulting.fr",
}

# =============================================================================
# PAGE 1 – COUVERTURE
# =============================================================================
PAGE_01 = {
    "titre_1": "RAPPORT",
    "titre_2": "D'ACTIVITÉS",
    "titre_3": "PROFESSIONNELLES",
    "bts": "BTS Management Économique de la Construction",
    "session": "SESSION 2026",
    "candidat_info": "Candidat n° 02537399911 | Académie de Lyon",
}

# =============================================================================
# PAGE 2 – FICHE D'IDENTITÉ DU CANDIDAT
# =============================================================================
PAGE_02 = {
    "titre": "FICHE D'IDENTITÉ DU CANDIDAT",
    "champs": [
        ("Candidat", "BAHAFID Mohamed"),
        ("N° Candidat", "02537399911"),
        ("Académie", "Lyon"),
        ("Structure d'accueil", "Conseil Régional de Béni Mellal-Khénifra (Maroc)"),
        ("Direction", "Agence d'Exécution des Projets"),
        ("Poste occupé", "Technicien Études et Suivi des Travaux"),
        ("Durée d'expérience", "8 ans dans le BTP (3 ans Maroc + 5 ans France)"),
        ("Formation BIM", "Technicien Modeleur BIM – AFPA Colmar (8 mois)"),
        ("Activité actuelle", "BIMCO – Projeteur BIM / Économiste de la construction"),
        ("SIREN / APE", "999580053 / 7112B – Ingénierie, études techniques"),
        ("Siège", "44 rue de la République, 42510 Bussières (Loire)"),
    ],
    "pied": "gestion.bimco-consulting.fr | BIMCO – Expert BIM & Économie de la Construction",
}

# =============================================================================
# PAGE 3 – SOMMAIRE
# =============================================================================
PAGE_03 = {
    "titre": "SOMMAIRE",
    "sections": [
        ("01", "Introduction",
         "Parcours, contexte, objectifs du rapport", "p. 4"),
        ("02", "Cadre professionnel",
         "Mon parcours | Conseil Régional | BIMCO | Outils", "p. 5-12"),
        ("03", "Projet 1 : Mise à niveau 4 communes",
         "53,5 M DH TTC – 8 corps d'état – Province de Khénifra", "p. 13-21"),
        ("04", "Projet 2 : Route Lehri-Kerrouchen",
         "29 M DH TTC – 25 km en zone montagneuse", "p. 22-25"),
        ("05", "Activités complémentaires",
         "5 autres marchés au Maroc + Expérience terrain France", "p. 26"),
        ("06", "Compétences et analyse",
         "Référentiel BTS MEC + Comparaison Maroc / France", "p. 27-28"),
        ("07", "Projet professionnel",
         "Court, moyen et long terme – Vision BIMCO", "p. 29"),
        ("08", "Conclusion",
         "Bilan et perspectives", "p. 30"),
    ],
}

# =============================================================================
# PAGE 4 – INTRODUCTION
# =============================================================================
PAGE_04 = {
    "titre": "INTRODUCTION",
    "points": [
        ("8 ans d'expérience dans le BTP",
         "Parcours complet depuis la formation au Maroc jusqu'à la création de BIMCO en France"),
        ("Maîtrise d'ouvrage publique au Maroc (4,5 ans)",
         "7 marchés publics suivis au Conseil Régional de Béni Mellal-Khénifra, +100 M DH d'investissements"),
        ("Chef de chantier gros œuvre en France (5 ans)",
         "Coffrage, ferraillage, bétonnage chez Ergalis (Feurs) et Minssieux et Fils (Mornant)"),
        ("Formation BIM Modeleur – AFPA Colmar (8 mois)",
         "Revit, Navisworks, Dynamo, Python, C#/API Revit, formats IFC"),
        ("Création de BIMCO – Indépendant (janv. 2026)",
         "Projeteur BIM et Économiste de la construction – SIREN 999580053"),
    ],
    "projets_titre": "Ce rapport présente deux projets majeurs réalisés au Conseil Régional :",
    "projets": [
        "Projet 1 : Mise à niveau de 4 communes (53,5 M DH TTC, 8 corps d'état)",
        "Projet 2 : Construction route Lehri-Kerrouchen, 25 km (29 M DH TTC)",
    ],
    "citation": "« Cette double expérience, côté maîtrise d'ouvrage publique et côté\nexécution, m'a offert une vision complète du cycle de vie d'un projet. »",
    "etapes": ["Formation\nMaroc", "MOA\nMaroc", "Chantier\nFrance", "BIM\nColmar", "BIMCO"],
    "dates_etapes": ["2014-17", "2017-22", "2022-24", "2023-24", "2026"],
}

# =============================================================================
# PAGE 5 – SÉPARATEUR PARTIE 1
# =============================================================================
PAGE_05 = {
    "numero": "01",
    "titre": "CADRE PROFESSIONNEL",
    "sous_titre": "Mon parcours en 5 phases | Le Conseil Régional de Béni Mellal-Khénifra\nBIMCO – Mon activité indépendante | Outils et méthodes",
}

# =============================================================================
# PAGE 6 – MON PARCOURS (TIMELINE)
# =============================================================================
PAGE_06 = {
    "titre": "MON PARCOURS EN 5 PHASES",
    "phases": [
        ("2014-2017", "Formation BTP au Maroc",
         "Technicien Chef de Chantier BTP\nDiplôme TS Gros Œuvre à l'ISBTP\nPlanification, organisation, métrés"),
        ("2017-2022", "Chargé d'affaires – MOA",
         "Conseil Régional Béni Mellal-Khénifra\n7 marchés publics, +100 M DH\nEstimations, DCE, analyse offres, suivi"),
        ("2022-2024", "Chef de chantier France",
         "Ergalis (Feurs, 42) puis Minssieux (69)\nCoffrage, banches, ferraillage, béton\nGestion d'équipe, contrôle qualité"),
        ("2023-2024", "Formation BIM – AFPA",
         "8 mois à Colmar (68)\nRevit, Navisworks, Dynamo, C#/API\nIFC, extraction quantités, maquettes"),
        ("Depuis 2026", "Création BIMCO",
         "Micro-entreprise, APE 7112B\nProjeteur BIM + Économiste construction\nApp « Gestion Chantiers » développée"),
    ],
}

# =============================================================================
# PAGE 7 – CHIFFRES CLÉS
# =============================================================================
PAGE_07 = {
    "titre": "MON PARCOURS EN CHIFFRES",
    "chiffres": [
        ("8", "ANS", "d'expérience BTP",
         "Depuis la formation initiale\nau Maroc en 2014"),
        ("2", "PAYS", "Maroc + France",
         "MOA publique à Khénifra\npuis chantier en Rhône-Alpes"),
        ("7", "MARCHÉS", "publics suivis",
         "Routes, pistes, VRD, AEP,\naménagement urbain"),
        ("+100M", "DH", "d'investissements",
         "Volume global des marchés\ngérés (~9 M€)"),
        ("5", "ANS", "sur chantier",
         "Chef d'équipe et chef de\nchantier gros œuvre"),
        ("8", "MOIS", "de formation BIM",
         "AFPA Colmar : Revit,\nNavisworks, Dynamo, C#"),
        ("10+", "LOGICIELS", "maîtrisés",
         "BIM, DAO, bureautique,\ndéveloppement web"),
        ("1", "ENTREPRISE", "créée (BIMCO)",
         "Micro-entreprise depuis\njanvier 2026, APE 7112B"),
    ],
}

# =============================================================================
# PAGE 8 – CONSEIL RÉGIONAL
# =============================================================================
PAGE_08 = {
    "sur_titre": "STRUCTURE D'ACCUEIL",
    "titre": "Conseil Régional de\nBéni Mellal-Khénifra",
    "facts": [
        ("5 Provinces", "Béni Mellal, Azilal, Fquih Ben Salah,\nKhénifra, Khouribga"),
        ("28 374 km²", "Superficie de la région"),
        ("2,5 millions", "Habitants"),
        ("Compétences", "Routes régionales et communales,\naménagement urbain et rural,\nadduction d'eau potable,\néquipements publics"),
        ("Budget invest.", "Centaines de millions de DH/an\nconsacrés aux infrastructures"),
    ],
    "reglement_titre": "Décret n°2-12-349\ndu 20 mars 2013",
    "reglement_desc": "Cadre réglementaire marchés publics\nAppel d'offres ouvert, restreint\net marché négocié",
}

# =============================================================================
# PAGE 9 – ORGANIGRAMME
# =============================================================================
PAGE_09 = {
    "titre": "MON POSTE AU SEIN DE L'AGENCE",
    "sous_titre": "Agence d'Exécution des Projets",
    "postes": {
        "president": "Président du\nConseil Régional",
        "directeur": "Directeur de l'Agence\nM. DOGHMANI",
        "service_1": "Service Études\n& Programmation",
        "service_2": "Service Marchés\n& Contrats",
        "service_3": "Service Suivi\ndes Travaux",
        "mon_poste": "Technicien de Suivi\nBAHAFID Mohamed",
    },
    "label_poste": "MON POSTE",
    "missions": [
        "Missions principales :",
        "- Métrés détaillés et estimations de l'administration",
        "- Préparation des dossiers de consultation (CPS, RC, BPDE)",
        "- Analyse des offres et participation à la CAO",
        "- Suivi financier : situations, décomptes, avenants",
        "- Visites de chantier et contrôle de conformité",
    ],
}

# =============================================================================
# PAGE 10 – BIMCO
# =============================================================================
PAGE_10 = {
    "nom": "BIMCO",
    "slogan": "Expert BIM & Économie de la Construction",
    "infos": [
        "Créé le 9 janvier 2026 | Micro-entreprise",
        "SIREN : 999580053 | SIRET : 99958005300018",
        "Code APE : 7112B – Ingénierie, études techniques",
        "44 rue de la République, 42510 Bussières (Loire)",
        "BIM + COnstruction = BIMCO",
    ],
    "domaines": [
        ("Projeteur\nBIM", "Revit, Navisworks"),
        ("Métrés\nQuantitatifs", "BIM + traditionnel"),
        ("Études\nde prix", "Estimations, DPGF"),
        ("Outils\nnumériques", "Scripts, plugins"),
        ("Suivi\néconomique", "Budgets, décomptes"),
    ],
}

# =============================================================================
# PAGE 11 – OUTILS ET MÉTHODES
# =============================================================================
PAGE_11 = {
    "titre": "NOS OUTILS ET MÉTHODES",
    "outils": [
        ("Revit", "Modélisation BIM Architecture + Structure, extraction de quantités", 90, "Expert"),
        ("Navisworks", "Coordination et synthèse de maquettes, détection de clashs", 85, "Expert"),
        ("Dynamo", "Scripts visuels pour automatiser les workflows Revit", 80, "Autonome"),
        ("Python", "Automatisation de tâches BTP, traitement de données, scripts", 80, "Autonome"),
        ("C# / API Revit", "Développement de plugins Revit personnalisés", 75, "Autonome"),
        ("Excel avancé", "Métrés, estimations, bases de prix, tableaux de suivi financier", 95, "Expert"),
        ("AutoCAD", "Lecture et exploitation de plans, pièces dessinées", 85, "Expert"),
        ("MS Project", "Planification de travaux, suivi d'avancement", 70, "Autonome"),
    ],
    "dev_titre": "Développement d'outils numériques pour le BTP",
    "dev_stack": "React/TypeScript | Node.js/Express | PostgreSQL | Electron | Docker\nApplication « Gestion Chantiers » développée pour BIMCO",
    "protocole_bim": "Convention BIM : LOD 300 / LOI 3 | Échange IFC 2x3 | Plateforme collaborative\nWorkflow : Modélisation Revit → Export IFC → Extraction quantités → Chiffrage DPGF → Reporting",
    "pied": "Format IFC | Interopérabilité BIM | Open BIM | Extraction automatique de quantités",
}

# =============================================================================
# PAGE 12 – APPLICATION GESTION CHANTIERS
# =============================================================================
PAGE_12 = {
    "sur_titre": "RÉALISATION PHARE",
    "titre": "Application « Gestion Chantiers »",
    "url": "gestion.bimco-consulting.fr",
    "placeholder": "[CAPTURES D'ÉCRAN DE L'APPLICATION\nÀ FOURNIR PAR LE CANDIDAT]",
    "modules": [
        ("Devis", "Bibliothèque d'ouvrages,\nimport/export Excel"),
        ("Chantiers", "Multi-projets, cartes,\nKanban, budget 12 postes"),
        ("Facturation", "Situations de travaux,\navancement, OCR, relances"),
        ("Équipes", "Planning hebdo,\naffectations, compétences"),
        ("Main d'œuvre", "Pointage, planification,\nvariables de paie"),
        ("Finances", "30+ indicateurs,\nrentabilité, tableaux de bord"),
        ("Documents", "GED, fiches intervention\nphoto + signature"),
        ("Appro.", "Fournisseurs, catalogue,\ncommandes, stocks"),
    ],
    "stack": "Stack technique : React/TypeScript | Node.js/Express | PostgreSQL | Electron | Docker",
    "deploiement": "Déployé sur serveur NAS via Docker | Version web + application desktop Windows",
}

# =============================================================================
# PAGE 13 – SÉPARATEUR PARTIE 2
# =============================================================================
PAGE_13 = {
    "numero": "02",
    "titre": "ACTIVITÉS ET PROJETS RÉALISÉS",
    "sous_titre": "Projet 1 : Mise à niveau 4 communes (53,5 M DH – 8 corps d'état)\nProjet 2 : Route Lehri-Kerrouchen 25 km (29 M DH – zone montagneuse)",
}

# =============================================================================
# PAGE 14 – PROJET 1 : FICHE D'IDENTITÉ
# =============================================================================
PAGE_14 = {
    "sur_titre": "PROJET 1",
    "titre": "Mise à niveau des centres\nde 4 communes",
    "fiche": [
        ("Marché", "n°38-RBK-2017 (Lot 4)"),
        ("Maître d'ouvrage", "Conseil Régional BMK"),
        ("Localisation", "Province de Khénifra"),
        ("4 communes", "El Hammam, Kerrouchen,\nOuaoumana, Sebt Ait Rahou"),
        ("Nature", "Aménagement urbain - VRD"),
        ("8 parties", "Assainissement, chaussée, trottoirs,\nsignalisation, éclairage, murs,\npaysager, mobilier urbain"),
    ],
    "montant": "53,5 M DH TTC",
    "detail_montant": "soit 4,86 M€ | 44,6 M DH HT + TVA 20% | Lot unique | 4 communes",
    "pied": "Pièces du marché : CPS + RC + BPDE + Plans | Appel d'offres ouvert",
}

# =============================================================================
# PAGE 15 – RÉPARTITION BUDGÉTAIRE PAR COMMUNE
# =============================================================================
PAGE_15 = {
    "titre": "RÉPARTITION BUDGÉTAIRE PAR COMMUNE",
    "communes": [
        ("El Hammam", "6,6 M DH HT", "14,9%", "7,96 M DH TTC"),
        ("Kerrouchen", "7,3 M DH HT", "16,5%", "8,80 M DH TTC"),
        ("Ouaoumana", "15,8 M DH HT", "35,5%", "19,0 M DH TTC"),
        ("Sebt Ait Rahou", "14,8 M DH HT", "33,2%", "17,8 M DH TTC"),
    ],
    "callout": "68% du budget concentré sur 2 communes\nOuaoumana + Sebt Ait Rahou = 30,6 M DH HT",
}

# =============================================================================
# PAGE 16 – 8 CORPS D'ÉTAT
# =============================================================================
PAGE_16 = {
    "titre": "8 CORPS D'ÉTAT",
    "parties": [
        ("01", "Assainissement", "Tranchées, buses PEHD/PVC,\nregards, bouches d'égout"),
        ("02", "Chaussée", "Terrassement, GNF, GNA,\nenrobés bitumineux"),
        ("03", "Trottoirs", "Bordures T1/T3, carreaux\nstriés, pavés, béton"),
        ("04", "Signalisation", "Marquage au sol, panneaux,\npeinture bordures"),
        ("05", "Éclairage public", "Tranchées, tubes annelés,\nmassifs candélabres, câbles"),
        ("06", "Murs et ouvrages", "Béton armé, maçonnerie\nmoellons, gabions"),
        ("07", "Aménagement paysager", "Terre végétale, réseau\nd'arrosage, plantation"),
        ("08", "Mobilier urbain", "Corbeilles, bancs\nen granite"),
    ],
}

# =============================================================================
# PAGE 17 – MES MISSIONS SUR LE PROJET 1
# =============================================================================
PAGE_17 = {
    "titre": "MES MISSIONS SUR LE PROJET 1",
    "missions": [
        ("A", "Estimation de\nl'administration",
         "Métrés détaillés des 8 parties\npour 4 communes\nSurfaces, linéaires, cubatures\nPrix unitaires de référence\nTableurs Excel structurés"),
        ("B", "Préparation\ndu DCE",
         "Rédaction du CPS (clauses\nadmin. et techniques)\nRC, BPDE, plans et pièces\nCohérence CPS / BPDE\nNomenclature des prix"),
        ("C", "Analyse des\noffres",
         "Vérification arithmétique\nDétection prix anormaux\nTableau comparatif des offres\nGrille de notation (/100)\nParticipation à la CAO"),
        ("D", "Suivi\nfinancier",
         "Situations de travaux\nmensuelles vérifiées\nDécomptes provisoires\nGestion des avenants\nTableau de bord par commune"),
    ],
    "pied": [
        "Cycle complet : de l'estimation à la réception provisoire (PV signé) et levée des réserves",
        "Réunions de chantier hebdomadaires | Suivi des approvisionnements (enrobés, buses, candélabres)",
        "Communication : reporting au Directeur, comptes rendus écrits, confidentialité des offres CAO",
    ],
}

# =============================================================================
# PAGE 18 – MÉTRÉS OUAOUMANA
# =============================================================================
PAGE_18 = {
    "titre": "FOCUS : MÉTRÉS – COMMUNE D'OUAOUMANA",
    "chiffres": [
        ("3 620 m³", "Tranchées"),
        ("2 370 ml", "Buses PEHD"),
        ("56", "Regards de visite"),
        ("2 175 T", "Enrobés"),
        ("11 500 m²", "Carreaux striés"),
        ("163", "Massifs candélabres"),
        ("4 800 ml", "Bordures T3"),
    ],
    "methode": "Méthode : métrés sur plans du bureau d'études, complétés par relevés sur site.\nTableurs Excel structurés avec formules de calcul vérifiables.",
    "citation": "« La précision des avant-métrés est fondamentale :\nune erreur de 5% sur les trottoirs peut représenter +1 million de DH d'écart. »",
}

# =============================================================================
# PAGE 19 – ANALYSE DES OFFRES
# =============================================================================
PAGE_19 = {
    "titre": "ANALYSE DES OFFRES & CAO",
    "resultat": "Entreprise B retenue avec 94/100",
    "criteres": "Critères : offre financière (70 pts) + moyens humains (10) + matériels (10) + références (10)",
    "docs_titre": "Documents officiels de la procédure",
    "legende_cao": "Convocation officielle de la\nCommission d'Appel d'Offres\nRégion Béni Mellal-Khénifra",
    "legende_rejet": "Notification de rejet d'offre\nProcédure de marchés publics\nmarocains (Décret n°2-12-349)",
    # Données du graphique
    "notation": {
        "criteres": ['Offre financière\n(/70)', 'Moyens humains\n(/10)', 'Moyens matériels\n(/10)', 'Références\n(/10)'],
        "entreprise_a": (62, 7, 8, 6),
        "entreprise_b": (68, 9, 9, 8),
        "entreprise_c": (55, 6, 7, 5),
    },
}

# =============================================================================
# PAGE 20 – SUIVI FINANCIER KERROUCHEN
# =============================================================================
PAGE_20 = {
    "titre": "SUIVI FINANCIER – KERROUCHEN",
    "chiffres": [
        ("Marché", "7 336 914 DH HT"),
        ("Exécuté", "7 394 495 DH HT"),
        ("Écart", "+57 581 DH (+0,8%)"),
    ],
    "callout_titre": "DÉPASSEMENT MAÎTRISÉ : +0,8%",
    "callout_detail": "Aucun avenant nécessaire | Écarts dans la tolérance contractuelle\nPostes Chaussée et Murs en plus-value (terrain rocheux)\nPoste Paysager réduit pour compenser partiellement",
}

# =============================================================================
# PAGE 21 – DIFFICULTÉS ET SOLUTIONS PROJET 1
# =============================================================================
PAGE_21 = {
    "titre": "DIFFICULTÉS ET SOLUTIONS",
    "defis": [
        ("8 parties techniques",
         "Coordination complexe entre\nassainissement, chaussée, trottoirs,\nsignalisation, éclairage, murs,\npaysager et mobilier urbain",
         "Tableau de bord Excel consolidé\npermettant le suivi en temps réel\npar commune et par partie"),
        ("Conditions géologiques",
         "Terrain rocheux imprévu nécessitant\ndes plus-values de terrassement,\nnature du sol variable entre communes",
         "Visites de chantier régulières\npour valider les quantités déclarées\npar l'entreprise"),
        ("4 chantiers simultanés",
         "Gestion de 4 sites répartis sur\nla province de Khénifra,\nsuivi dispersé géographiquement",
         "Anticipation des avenants par\nveille continue sur les écarts\nde quantités"),
        ("Écarts quantités",
         "Différences entre les quantités\nestimées au BPDE et les quantités\nréellement exécutées sur chantier",
         "Communication régulière avec\nla hiérarchie sur l'état\nd'avancement financier"),
    ],
}

# =============================================================================
# PAGE 22 – PROJET 2 : FICHE D'IDENTITÉ
# =============================================================================
PAGE_22 = {
    "sur_titre": "PROJET 2",
    "titre": "ROUTE\nLEHRI-KERROUCHEN",
    "km": "25 KM",
    "montant": "29 M DH",
    "fiche_items": [
        "Marché n°46-RBK-2017 – Programme National Routes Rurales (PRR3)",
        "Désenclavement des zones rurales de la province de Khénifra",
        "Relief montagneux, 120 334 m³ de déblais, ouvrages hydrauliques",
        "3 sections : Linéaire 23 prix + Carrefour PK 0+000 (11 prix) + Bretelles (19 prix)",
        "53 prix unitaires au bordereau – Corps de chaussée + Ouvrages de traversée",
    ],
}

# =============================================================================
# PAGE 23 – MÉTRÉS ROUTIERS
# =============================================================================
PAGE_23 = {
    "titre": "PRINCIPAUX MÉTRÉS DE LA SECTION LINÉAIRE",
    "sous_titre": "Calculs à partir des profils en travers du bureau d'études routier + étude hydrologique",
    "metres": [
        ("120 334 m³", "Déblais"),
        ("76 735 m³", "Remblais"),
        ("19 989 m³", "Couche de base GNB"),
        ("34 481 m³", "Couche de fondation GNF2"),
        ("794 ml", "Buses Ø1000"),
        ("64,5 T", "Acier HA"),
        ("733 m³", "Béton B2"),
        ("4 579 m³", "Béton B3"),
        ("789 m³", "Gabions"),
    ],
}

# =============================================================================
# PAGE 24 – BUDGET ROUTE
# =============================================================================
PAGE_24 = {
    "titre": "RÉPARTITION BUDGÉTAIRE – PROJET 2",
    "items": [
        ("Terrassement", "4,3 M DH", "17,9%"),
        ("Corps de chaussée", "7,3 M DH", "30,1%"),
        ("Revêtement", "2,6 M DH", "10,7%"),
        ("Ouvrages hydrauliques", "6,8 M DH", "28,1%"),
        ("Soutènement", "0,46 M DH", "1,9%"),
        ("Bretelles + carrefour", "2,7 M DH", "11,3%"),
    ],
    "callout": "58% du budget sur 2 postes :\nCorps de chaussée + Ouvrages hydrauliques",
    "pied": "Caractéristique d'un projet routier en zone de montagne :\nnombreux talwegs traversés, ouvrages hydrauliques fréquents, terrain accidenté",
}

# =============================================================================
# PAGE 25 – DÉFIS CHANTIER ROUTIER
# =============================================================================
PAGE_25 = {
    "titre": "DÉFIS D'UN CHANTIER ROUTIER\nEN ZONE MONTAGNEUSE",
    "defis": [
        ("Relief montagneux",
         "Le tracé de 25 km traverse un\nterrain très accidenté, générant\n120 334 m³ de déblais et nécessitant\ndes gabions de soutènement (789 m³)"),
        ("Conditions climatiques",
         "Intempéries hivernales imposant\ndes arrêts de chantier : ordres de\nservice d'arrêt et de reprise émis\npour gel et fortes précipitations"),
        ("Écarts de cubatures",
         "Les cubatures de terrassement\nréelles se sont écartées des\nprévisions en raison de la géologie\nrencontrée (terrain rocheux imprévu)"),
        ("Ouvrages hydrauliques",
         "Le nombre et dimensionnement des\nouvrages de traversée ont été ajustés\naprès observation des crues pendant\nles travaux (buses, dalots, béton)"),
    ],
}

# =============================================================================
# PAGE 26 – ACTIVITÉS COMPLÉMENTAIRES
# =============================================================================
PAGE_26 = {
    "titre": "ACTIVITÉS COMPLÉMENTAIRES",
    "maroc_titre": "5 AUTRES MARCHÉS AU CONSEIL RÉGIONAL",
    "marches": [
        ("27-RBK-2017", "Construction route village Sidi Bouabbad à Oued Grou",
         "Route / Piste rurale (CT Sidi Lamine)"),
        ("28-RBK-2017", "Route Ajdir-Ayoun Oum Errabia + Piste Lijon Kichchon",
         "Route + Piste (CT Sidi Lamine)"),
        ("30-RBK-2017", "Adduction en Eau Potable El Borj – El Hamam",
         "AEP (réseau d'eau potable)"),
        ("39-RBK-2017", "Pistes Hartaf – Sebt Ait Rahou",
         "Pistes rurales"),
        ("49-RBK-2016", "Aménagement voie Amghass – Bouchbel",
         "Voirie / Aménagement"),
    ],
    "france_titre": "EXPÉRIENCE TERRAIN EN FRANCE",
    "france": [
        ("2022-2023", "Chef d'équipe GO – Ergalis, Feurs (42)",
         "Lecture de plans, implantation et traçage, montage des banches,\nmise en place des armatures, coulage du béton, suivi des cycles de coffrage"),
        ("2024", "Chef de chantier – Minssieux et Fils, Mornant (69)",
         "Encadrement opérationnel d'équipe, organisation quotidienne,\nsuivi de l'avancement des travaux, contrôle qualité d'exécution"),
    ],
    "pied": "Apports pour l'économie de la construction : connaissance des coûts réels de production,\nmaîtrise des techniques de gros œuvre (coffrage, ferraillage, bétonnage), gestion d'équipes et planning",
}

# =============================================================================
# PAGE 27 – COMPÉTENCES BTS MEC
# =============================================================================
PAGE_27 = {
    "titre": "COMPÉTENCES BTS MEC ACQUISES",
    "competences": [
        ("Réaliser des métrés", 85, "Autonome", "7 marchés + extraction BIM"),
        ("Estimer un ouvrage", 80, "Autonome", "Estimations admin. + bordereaux"),
        ("Analyser des offres", 85, "Autonome", "Grilles notation, comparatifs, CAO"),
        ("Suivre un budget", 90, "Expert", "Décomptes, situations, 7 marchés"),
        ("Suivre un chantier", 95, "Expert", "5 ans terrain + 4,5 ans MOA"),
        ("Communiquer (C19)", 80, "Autonome", "Réunions, CR, reporting, CAO"),
        ("Rédiger pièces marchés", 80, "Autonome", "CPS, RC, BPDE, OS, PV"),
        ("Collaborer en BIM (T23)", 75, "Autonome", "IFC, convention BIM, LOD"),
        ("Modéliser en BIM", 85, "Autonome", "Revit, Navisworks, IFC"),
        ("Réglementation marchés", 60, "En progression", "Décret marocain + Code FR"),
    ],
    "distinction": "COMPÉTENCE DISTINCTIVE : TRIPLE VISION MOA / EXÉCUTION / BIM",
    "badges": [
        ("MOA", "Maîtrise d'ouvrage\npublique (4,5 ans)"),
        ("EXÉCUTION", "Chantier gros\nœuvre (5 ans)"),
        ("BIM", "Modélisation &\nAutomatisation (8 mois)"),
    ],
}

# =============================================================================
# PAGE 28 – COMPARAISON MAROC / FRANCE
# =============================================================================
PAGE_28 = {
    "sur_titre": "ANALYSE COMPARATIVE",
    "titre": "MAROC vs FRANCE",
    "comparaison": [
        ("Réglementation", "Décret n°2-12-349\ndu 20/03/2013", "Code de la commande\npublique"),
        ("Pièces du marché", "CPS + RC + BPDE\n+ Plans", "CCAP + CCTP + BPU/DQE\nou DPGF"),
        ("Normes", "Normes marocaines\nRPS 2000 (sismique)", "Eurocodes, DTU\nRE2020 (environnement)"),
        ("Suivi financier", "Attachements\ncontradictoires,\ndécomptes", "Situations de travaux\nmensuelles,\nrévision prix CCAP"),
        ("Passation", "Appel d'offres ouvert\nou restreint,\nmarché négocié", "Procédure formalisée\nou adaptée selon\nseuils européens"),
    ],
    "synthese": "Point commun : mêmes principes fondamentaux de la commande publique\nTransparence | Égalité de traitement | Mise en concurrence\nChoix de l'offre économiquement la plus avantageuse",
}

# =============================================================================
# PAGE 29 – PROJET PROFESSIONNEL
# =============================================================================
PAGE_29 = {
    "titre": "MON PROJET PROFESSIONNEL",
    "horizons": [
        ("COURT TERME", "2026",
         "Obtenir le BTS MEC\nLancer BIMCO (métrés, études\nde prix, suivi financier)\nDévelopper scripts et plugins\nRevit pour l'extraction\nautomatique de quantités"),
        ("MOYEN TERME", "2027-28",
         "Créer une gamme d'outils BIM\ndédiés au MEC :\n- Plugins extraction quantités\n- Chiffrage assisté maquette\n- Bases de prix connectées\n- Apps web suivi économique"),
        ("LONG TERME", "2029+",
         "BIMCO = cabinet d'ingénierie\nspécialisé BIM + Économie :\n- Prestations d'ingénierie\n  (métrés BIM, AMO, études)\n- Édition d'outils numériques\n  pour économistes"),
    ],
    "citation": "« Les outils numériques doivent être\nau service de l'économiste de la construction,\net non l'inverse. »",
    "avantage": "Ma double compétence – économiste formé au terrain ET développeur\nmaîtrisant le BIM – constitue un avantage différenciant rare dans le secteur.",
}

# =============================================================================
# PAGE 30 – CONCLUSION
# =============================================================================
PAGE_30 = {
    "titre": "CONCLUSION",
    "points": [
        ("8 ans de BTP entre Maroc et France",
         "De la formation initiale à la création de BIMCO, un parcours riche et complémentaire"),
        ("MOA publique + Exécution chantier + BIM",
         "Triple compétence rare : comprendre le maître d'ouvrage, le terrain et les outils numériques"),
        ("82,5 M DH de projets détaillés (7,4 M€)",
         "Mise à niveau de 4 communes (53,5 M DH) + Route Lehri-Kerrouchen 25 km (29 M DH)"),
        ("BIMCO : BIM + Économie de la Construction",
         "Des outils numériques au service de l'économiste : métrés, études de prix, suivi financier"),
    ],
    "citation": "« Le BTS MEC représente bien plus qu'un diplôme :\nc'est la validation officielle d'un parcours professionnel engagé\net le socle sur lequel je construirai des outils qui transformeront\nla pratique quotidienne de l'économie de la construction. »",
    "pied": "BAHAFID Mohamed | BIMCO | BTS MEC Session 2026 | Académie de Lyon",
}

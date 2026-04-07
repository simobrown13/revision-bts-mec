# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════╗
║  CONTENU DU RAPPORT U62 v9 – BAHAFID Mohamed               ║
║  BTS Management Économique de la Construction – Session 2026 ║
╠══════════════════════════════════════════════════════════════╣
║  VERSION 9 : Restructuration CPAR                           ║
║  - 5 situations professionnelles analysées (CPAR)           ║
║  - Tableau synthèse activités → compétences                 ║
║  - Bilan réflexif + Protocole BIM                           ║
║  - Pages fusionnées pour libérer l'espace                   ║
╚══════════════════════════════════════════════════════════════╝
"""

# =============================================================================
# INFORMATIONS CANDIDAT
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
# PAGE 1 – COUVERTURE (inchangée)
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
# PAGE 2 – FICHE D'IDENTITÉ DU CANDIDAT (inchangée)
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
# PAGE 3 – SOMMAIRE (mis à jour)
# =============================================================================
PAGE_03 = {
    "titre": "SOMMAIRE",
    "sections": [
        # "—" = non numéroté (intro/conclusion) | "01"-"04" = partie avec slide séparateur
        ("—", "Introduction et parcours",
         "Présentation du rapport, démarche CPAR, parcours en 5 phases", "p. 4"),
        ("01", "Cadre professionnel",
         "Conseil Régional de Béni Mellal-Khénifra | BIMCO", "p. 5–7"),
        ("02", "Projet 1 : Mise à niveau de 4 communes",
         "53,5 M DH TTC — Fiche projet, budget, 4 situations CPAR, défis", "p. 8–15"),
        ("03", "Projet 2 : Route Lehri-Kerrouchen",
         "29 M DH TTC — Fiche projet, budget, cubatures, défis montagne", "p. 16–20"),
        ("04", "Bilan, analyse et perspectives",
         "Activités complémentaires | Compétences | Réflexif | BIM | Projet pro", "p. 21–27"),
        ("—", "Conclusion et annexes",
         "Synthèse finale, documents officiels et photos de chantier", "p. 28–30"),
    ],
}

# =============================================================================
# PAGE 4 – INTRODUCTION + PARCOURS (fusion p.4+6+7)
# =============================================================================
PAGE_04 = {
    "titre": "INTRODUCTION ET PARCOURS",
    "intro_texte": (
        "Fort de huit années d'expérience dans le BTP — trois ans au Maroc en maîtrise d'ouvrage publique "
        "et cinq ans en France sur chantier et en formation BIM — ce rapport présente cinq situations "
        "professionnelles issues de deux projets majeurs : la mise à niveau de 4 communes (53,5 M DH TTC, "
        "8 corps d'état) et 25 km de route en montagne (29 M DH TTC). "
        "Cette double expérience MOA/exécution offre une vision complète du cycle de vie d'un projet : "
        "de la programmation jusqu'à la réception. Chaque situation est analysée selon la démarche CPAR."
    ),
    "phases": [
        ("2014-2017", "Formation BTP – Maroc",
         "CC BTP + TS Gros Œuvre ISBTP Oujda : planification, métrés, lecture de plans, suivi technique"),
        ("2017-2022", "Chargé d'affaires – MOA",
         "Conseil Régional BMK : 7 marchés (100 M DH) — avant-métrés, CAO, OS, décomptes définitifs"),
        ("2022-2024", "Chef de chantier – France",
         "Ergalis Feurs (banches/béton GO) + Minssieux Mornant (encadrement équipe) — coûts réels terrain"),
        ("2023-2024", "Formation BIM – AFPA Colmar",
         "Revit, Navisworks, Dynamo, C#/API, Python, IFC — 78 postes extraits, écart maquette/métré 1,8%"),
        ("Depuis 2026", "BIMCO – Indépendant",
         "Modélisation 3D, métrés BIM + traditionnel, études de prix, plugins Revit/Dynamo — APE 7112B"),
    ],
    "projets": [
        "Projet 1 : Mise à niveau de 4 communes — 53,5 M DH TTC",
        "Projet 2 : Route Lehri-Kerrouchen — 29 M DH TTC, 25 km",
    ],
    "chiffres_cles": [
        ("3,2%", "écart estimation P1"),
        ("94/100", "note CAO"),
        ("+0,8%", "dépassement maîtrisé"),
        ("0 avt", "avenant cubatures"),
    ],
}

# =============================================================================
# PAGE 5 – SÉPARATEUR "Cadre professionnel"
# =============================================================================
PAGE_05 = {
    "numero": "01",
    "titre": "CADRE PROFESSIONNEL",
    "sous_titre": "Le Conseil Régional de Béni Mellal-Khénifra | Mon poste\nBIMCO – Mon activité indépendante | Outils et méthodes",
}

# =============================================================================
# PAGE 6 – CONSEIL RÉGIONAL + MON POSTE (fusion p.8+9)
# =============================================================================
PAGE_06 = {
    "sur_titre": "STRUCTURE D'ACCUEIL",
    "titre": "Conseil Régional de Béni Mellal-Khénifra",
    "facts": [
        ("5 Provinces", "Béni Mellal, Azilal, Fquih Ben Salah, Khénifra, Khouribga"),
        ("28 374 km²", "Superficie de la région | 2,5 millions d'habitants"),
        ("Compétences", "Routes, aménagement urbain/rural, AEP, équipements publics"),
        ("Décret n°2-12-349", "Cadre réglementaire des marchés publics marocains"),
    ],
    "organigramme": {
        "president": "Président du Conseil Régional",
        "directeur": "Directeur de l'Agence – M. DOGHMANI",
        "services": ["Études &\nProgrammation", "Marchés &\nContrats", "Suivi des\nTravaux"],
        "mon_poste": "Technicien de Suivi\nBAHAFID Mohamed",
    },
    "missions": [
        "Métrés avant-projet (112 à 448 prix/commune) et estimation confidentielle de l'administration avant chaque AO — référence pour l'analyse des offres",
        "Rédaction des DCE : CPS, RC, BPDE avec vérification de cohérence des pièces écrites et des plans — zéro erreur contestée en commission",
        "Analyse des offres en CAO : grille de notation /100 (technique 60 pts + financier 40 pts), rapport d'analyse, procès-verbal de séance",
        "Suivi financier mensuel : tableau de bord à 3 indicateurs, vérification des situations mensuelles, décomptes provisoires et définitifs",
        "Visites hebdomadaires, attachements contradictoires, ordres de service d'arrêt/reprise, procès-verbaux de réception provisoire et définitive",
    ],
}

# =============================================================================
# PAGE 7 – BIMCO + OUTILS + APP (fusion p.10+11+12)
# =============================================================================
PAGE_07 = {
    "nom": "BIMCO",
    "slogan": "Expert BIM & Économie de la Construction",
    "infos": [
        "Créé le 9 janvier 2026 | SIREN : 999580053 | APE : 7112B",
        "44 rue de la République, 42510 Bussières (Loire)",
    ],
    "domaines": [
        ("Projeteur BIM", "Revit Architecture + Structure, Navisworks, IFC 2x3, Dynamo"),
        ("Métrés tous corps d'état", "Extraction BIM (nomenclatures Revit) + métré traditionnel sur plans"),
        ("Études de prix", "Estimations confidentielles, DPGF, sous-détails de prix, bordereau"),
        ("Outils numériques", "Python (automatisation), C# (plugins Revit), Apps web (React/Node)"),
        ("Suivi économique", "Tableau de bord, situations de travaux, décomptes, révision de prix"),
    ],
    "outils_principaux": [
        ("Revit", 90), ("Excel avancé", 95), ("AutoCAD", 85),
        ("Navisworks", 85), ("Python", 80), ("Dynamo", 80),
    ],
    "app": {
        "titre": "Application « Gestion Chantiers »",
        "url": "gestion.bimco-consulting.fr",
        "stack": "React/TypeScript | Node.js | PostgreSQL | Docker | Vercel",
    },
}

# =============================================================================
# PAGE 8 – SÉPARATEUR "Projet 1 : 5 situations analysées"
# =============================================================================
PAGE_08 = {
    "numero": "02",
    "titre": "PROJET 1 : MISE À NIVEAU\nDE 4 COMMUNES",
    "sous_titre": "53,5 M DH TTC – 8 corps d'état – Province de Khénifra\n5 situations professionnelles analysées (CPAR)",
}

# =============================================================================
# PAGE 9 – FICHE PROJET 1 (= page 14 actuelle)
# =============================================================================
PAGE_09 = {
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
# PAGE 10 – BUDGET + 8 CORPS D'ÉTAT (fusion p.15+16)
# =============================================================================
PAGE_10 = {
    "titre": "BUDGET ET 8 CORPS D'ÉTAT",
    "communes": [
        ("El Hammam", "6,6 M DH", "14,9%"),
        ("Kerrouchen", "7,3 M DH", "16,5%"),
        ("Ouaoumana", "15,8 M DH", "35,5%"),
        ("Sebt Ait Rahou", "14,8 M DH", "33,2%"),
    ],
    "parties": [
        ("01", "Assainissement", "Réseau eaux pluviales et usées, regards, collecteurs"),
        ("02", "Chaussée", "Terrassement, couches de base et de roulement"),
        ("03", "Trottoirs", "Bordures, dallage, revêtement piétonnier"),
        ("04", "Signalisation", "Marquage au sol, panneaux directionnels et réglementaires"),
        ("05", "Éclairage", "Candélabres, luminaires LED, réseau électrique"),
        ("06", "Murs", "Soutènement béton armé, maçonnerie, gabions"),
        ("07", "Paysager", "Plantation, engazonnement, arrosage intégré"),
        ("08", "Mobilier", "Corbeilles, bancs, bornes d'aménagement urbain"),
    ],
}

# =============================================================================
# PAGE 11 – SITUATION 1 : Estimation de l'administration
# =============================================================================
SITUATION_1 = {
    "numero": 1,
    "titre": "Estimation confidentielle de l'administration – Ouaoumana",
    "chiffre_cle": "15,8 M DH HT",
    "chiffre_label": "Estimation validée — écart de 3,2% avec l'offre retenue",
    "competence": "Réaliser des métrés tous corps d'état et estimer un ouvrage",
    "contexte": (
        "Pour le marché de mise à niveau d'Ouaoumana, j'ai été chargé "
        "d'établir l'estimation confidentielle de l'administration — le "
        "document qui fixe le prix plafond avant l'appel d'offres. "
        "Le périmètre : 15,8 M DH HT, soit 35% du budget global du "
        "programme. Cette estimation devait être fiable, défendable et "
        "livrée en trois semaines pour respecter la date de publication."
    ),
    "probleme": (
        "Deux obstacles majeurs. D'abord, les plans d'assainissement "
        "du bureau d'études présentaient des incohérences — profils en "
        "long non calés, regards mal positionnés — ce qui faussait les "
        "métrés. Ensuite, la mercuriale de référence datait de 2014 : "
        "les prix des enrobés, des canalisations et de l'acier avaient "
        "dérivé de 15 à 22%. Impossible de chiffrer correctement sans "
        "actualiser ces données."
    ),
    "action": (
        "J'ai structuré le travail en trois phases. Métrés d'abord : "
        "112 lignes de prix relevées sur AutoCAD, complétées par quatre "
        "visites de terrain pour caler les altimétries manquantes. "
        "Puis la construction des prix unitaires en croisant trois "
        "sources : la mercuriale actualisée, cinq marchés similaires "
        "récemment adjugés dans la région, et huit devis fournisseurs "
        "locaux. Enfin, une relecture croisée avec un collègue avant "
        "validation par le Directeur."
    ),
    "resultat": (
        "L'estimation a été livrée dans le délai de trois semaines. "
        "Après ouverture des plis, l'écart avec l'offre retenue n'était "
        "que de 3,2% — bien en dessous de la marge habituelle de 5 à "
        "10%. Le Directeur a adopté ma méthode de croisement des sources "
        "comme standard pour les trois autres communes du programme. "
        "Aucune contestation lors de la Commission d'Appel d'Offres."
    ),
}

# =============================================================================
# PAGE 12 – SITUATION 2 : Analyse des offres et CAO
# =============================================================================
SITUATION_2 = {
    "numero": 2,
    "titre": "Analyse des offres et Commission d'Appel d'Offres",
    "chiffre_cle": "94/100",
    "chiffre_label": "Note de l'entreprise retenue — zéro recours",
    "competence": "Analyser les offres et préparer la décision en commission",
    "contexte": (
        "Trois entreprises ont répondu à l'appel d'offres. En tant que "
        "membre technique de la commission, j'étais responsable de "
        "l'analyse comparative : vérifier la conformité administrative "
        "de chaque dossier, contrôler la cohérence des prix proposés, "
        "et rédiger le rapport d'analyse qui allait fonder la décision "
        "d'attribution. Un travail qui exige rigueur et impartialité."
    ),
    "probleme": (
        "Deux dossiers posaient problème. L'un contenait sept erreurs "
        "arithmétiques dans les sous-détails de prix. L'autre affichait "
        "des prix anormalement bas sur trois postes majeurs représentant "
        "42% du montant global — des écarts de 25 à 33% par rapport "
        "aux prix du marché. Des offres de ce type, si elles sont "
        "retenues, conduisent presque toujours à des litiges en phase "
        "d'exécution."
    ),
    "action": (
        "J'ai procédé méthodiquement. D'abord, la vérification de "
        "conformité des trois dossiers. Puis la correction arithmétique "
        "des erreurs — en appliquant la règle de primauté du prix "
        "unitaire sur le total. J'ai ensuite demandé par écrit à "
        "l'entreprise concernée de justifier ses prix bas ; sa réponse "
        "a été jugée insuffisante par la commission. Enfin, j'ai bâti "
        "une grille comparative sur les dix postes les plus lourds, "
        "avec une notation technique et financière sur 100 points."
    ),
    "resultat": (
        "L'entreprise la mieux notée a obtenu 94/100. L'attribution "
        "s'est faite en quinze jours, sans aucun recours déposé — signe "
        "que l'analyse était solide et transparente. Le Directeur a "
        "validé le rapport et le PV de commission sans réserve. Cette "
        "expérience m'a appris qu'une analyse rigoureuse en amont "
        "protège le maître d'ouvrage tout au long du marché."
    ),
}

# =============================================================================
# PAGE 13 – SITUATION 3 : Suivi financier Kerrouchen
# =============================================================================
SITUATION_3 = {
    "numero": 3,
    "titre": "Suivi financier et tableau de bord – Kerrouchen",
    "chiffre_cle": "+0,8%",
    "chiffre_label": "Dépassement final maîtrisé — avenant évité",
    "competence": "Suivre l'exécution financière et anticiper les dérives",
    "contexte": (
        "Le chantier de Kerrouchen représentait 7,3 M DH sur 18 mois. "
        "J'assurais le suivi mensuel complet : vérification des situations "
        "de travaux présentées par l'entreprise, contrôle des quantités "
        "réellement exécutées, et transmission des décomptes au Directeur. "
        "La découverte d'un terrain rocheux imprévu en cours de chantier "
        "a rendu ce suivi particulièrement critique."
    ),
    "probleme": (
        "À mi-parcours, deux postes dérapaient : la chaussée à +12% à "
        "cause du terrassement rocheux, et les murs à +15% en raison de "
        "fondations plus profondes que prévu. En extrapolant, le "
        "dépassement global atteignait +4,8%, soit 349 000 DH. Dépasser "
        "le seuil de 5% aurait imposé un avenant — trois à six mois de "
        "procédure administrative, avec un risque d'arrêt du chantier."
    ),
    "action": (
        "J'ai mis en place un tableau de bord hebdomadaire à trois "
        "indicateurs : avancement physique, consommation budgétaire et "
        "écart prévisionnel. Chaque poste critique faisait l'objet "
        "d'un attachement contradictoire sur le terrain. Pour absorber "
        "la dérive, j'ai proposé au Directeur un mécanisme de "
        "compensation : réduire les prestations paysagères et le "
        "mobilier urbain, libérant 56 000 DH de marge. Un arbitrage "
        "pragmatique, validé sans difficulté."
    ),
    "resultat": (
        "Dépassement final ramené à +0,8% — soit 292 000 DH sous le "
        "seuil d'avenant. Le chantier s'est terminé sans procédure "
        "supplémentaire. Mon tableau de bord a ensuite été répliqué "
        "sur les trois autres communes du programme et adopté comme "
        "outil de référence par l'Agence. C'est cette situation qui "
        "m'a convaincu de l'importance d'un pilotage financier en "
        "temps réel."
    ),
}

# =============================================================================
# PAGE 14 – SITUATION 4 : Communication multi-sites
# =============================================================================
SITUATION_4 = {
    "numero": 4,
    "titre": "Communication et coordination de 4 chantiers simultanés",
    "chiffre_cle": "4 sites",
    "chiffre_label": "Coordonnés en parallèle — 3 crises résolues en 48h",
    "competence": "Communiquer et coordonner en contexte multi-sites",
    "contexte": (
        "J'étais le relais unique du Directeur pour quatre chantiers VRD "
        "répartis sur la province de Khénifra, distants de 20 à 80 km. "
        "Au quotidien, je faisais l'interface entre l'entreprise, le "
        "bureau d'études, le laboratoire de contrôle et la hiérarchie "
        "régionale. Chaque décision devait être tracée par écrit — un "
        "réflexe qui s'est avéré déterminant."
    ),
    "probleme": (
        "La semaine 23, tout s'est accéléré. Trois crises simultanées : "
        "un retard de livraison des enrobés à Ouaoumana, une alerte "
        "météo nécessitant l'arrêt du chantier à Kerrouchen, et un "
        "litige sur les quantités de bordures à Sebt Ait Rahou — 15% "
        "d'écart entre ce que l'entreprise déclarait et ce qui était "
        "réellement posé. Le paiement mensuel était bloqué."
    ),
    "action": (
        "J'ai traité les trois fronts en 48 heures. Pour le retard "
        "d'enrobés, j'ai rédigé une note factuelle au Directeur avec "
        "un planning de rattrapage. Pour l'arrêt météo, j'ai émis "
        "l'ordre de service dès le lendemain avec photos datées à "
        "l'appui. Pour le litige bordures, je me suis déplacé sur "
        "place pour un re-mesurage contradictoire complet. En "
        "parallèle, un point quotidien à 8h avec les quatre chefs "
        "de chantier et un compte-rendu consolidé sous 24 heures."
    ),
    "resultat": (
        "Le retard enrobés a été rattrapé en deux semaines. L'arrêt "
        "chantier a été formalisé proprement, la reprise ordonnée. "
        "L'écart sur les bordures est tombé de 15% à 2,3%, débloquant "
        "le paiement. À la suite de cet épisode, le Directeur m'a "
        "confié la rédaction systématique de tous les ordres de service "
        "et comptes-rendus. Le modèle que j'avais créé a été adopté "
        "pour l'ensemble des marchés de l'Agence."
    ),
}

# =============================================================================
# PAGE 15 – DIFFICULTÉS ET SOLUTIONS P1 (= page 21 actuelle)
# =============================================================================
PAGE_15 = {
    "titre": "DIFFICULTÉS ET SOLUTIONS – PROJET 1",
    "defis": [
        ("8 parties techniques",
         "Les interdépendances entre corps d'état imposaient un séquençage rigoureux : "
         "assainissement terminé avant chaussée, éclairage avant trottoirs, murs de soutènement avant terrassement. "
         "Sur 4 communes simultanées, le moindre retard sur un poste créait un effet domino sur les suivants. "
         "La dispersion géographique (20 à 80 km) rendait le contrôle quotidien de chaque site impossible. "
         "32 lignes de suivi à tenir en parallèle — sans outil consolidé, les écarts auraient été invisibles.",
         "Tableau de bord Excel consolidé (8 parties × 4 communes) avec avancement physique et financier "
         "hebdomadaire. Jalons critiques identifiés par poste et par commune. Réunion de coordination "
         "mensuelle avec l'entreprise et le BET. CR transmis sous 24h au Directeur après chaque visite. "
         "Ce dispositif a permis de détecter les dérives dès le premier mois — et d'agir avant le seuil d'avenant."),
        ("Terrain rocheux imprévu",
         "L'étude géotechnique n'avait pas identifié le sous-sol réel. À Kerrouchen, dès le PK 4+200, "
         "le terrassement a révélé du calcaire fracturé — coût 85 DH/m³ vs 28 DH/m³ pour les déblais ordinaires. "
         "Surcoût immédiat sur la chaussée : +12%, soit 87 600 DH. Le seuil d'avenant étant à 5% du montant global, "
         "ce seul poste menaçait d'y contraindre l'Agence. "
         "Un avenant aurait signifié 3 à 6 mois de procédure administrative et un blocage des paiements.",
         "Attachements contradictoires à chaque découverte de terrain rocheux, avec photos datées géolocalisées. "
         "Reclassification documentée et formalisée. Mécanisme de compensation inter-postes proposé au Directeur : "
         "réduction des prestations paysagères (–44 000 DH) et du mobilier urbain (–12 000 DH), "
         "libérant 56 000 DH de marge. Résultat : dépassement de Kerrouchen ramené à +0,8% — sans avenant."),
        ("4 chantiers simultanés",
         "Un seul technicien pour le suivi de 4 chantiers distants de 20 à 80 km. "
         "Chaque site avait ses propres aléas, ses propres équipes, son propre avancement. "
         "En semaine 23 : 3 crises simultanées — retard de livraison des enrobés à Ouaoumana, "
         "alerte météo nécessitant un arrêt à Kerrouchen, litige sur les quantités de bordures "
         "à Sebt Ait Rahou (–15% d'écart contesté par l'entreprise). Le paiement mensuel était bloqué.",
         "Planning tournant de visites (1 site/jour, rotation hebdomadaire). Point téléphonique quotidien "
         "à 8h avec les 4 chefs de chantier. CR consolidé transmis sous 24h au Directeur. "
         "Semaine 23 : 3 fronts traités en 48h — note factuelle Directeur, OS d'arrêt avec photos datées, "
         "re-mesurage contradictoire sur place pour le litige bordures (écart ramené de 15% à 2,3%). "
         "Le modèle de CR créé dans l'urgence a ensuite été adopté comme standard par l'Agence."),
        ("Écarts de quantités",
         "Les différences entre le BPDE initial et les quantités réellement exécutées atteignaient "
         "+15% sur les bordures à Sebt Ait Rahou et +12% sur la chaussée à Kerrouchen. "
         "Ces écarts résultaient de plans imprécis et de conditions de terrain non anticipées. "
         "En extrapolant les dérives à mi-chantier, le dépassement global atteignait +4,8% — "
         "juste sous le seuil d'avenant de 5%, mais sans marge de sécurité. "
         "Un dépassement de ce seuil aurait bloqué les paiements pendant 3 à 6 mois.",
         "Veille continue via le tableau de bord hebdomadaire — tous les postes suivis, pas seulement les critiques. "
         "Identification des postes compensatoires (ceux exécutés en dessous du BPDE). "
         "Proposition formelle au Directeur : réduction Paysager (–44 000 DH) + Mobilier (–12 000 DH) "
         "= 56 000 DH de marge dégagée. Dépassement final de Kerrouchen : +0,8%. "
         "Méthode adoptée comme standard pour les 3 autres communes dès la S24."),
    ],
}

# =============================================================================
# PAGE 16 – SÉPARATEUR "Projet 2"
# =============================================================================
PAGE_16 = {
    "numero": "03",
    "titre": "PROJET 2 : ROUTE\nLEHRI-KERROUCHEN",
    "sous_titre": "29 M DH TTC – 25 km en zone montagneuse\nProgramme National des Routes Rurales (PRR3)",
}

# =============================================================================
# PAGE 17 – FICHE P2 + MÉTRÉS (fusion p.22+23)
# =============================================================================
PAGE_17 = {
    "sur_titre": "PROJET 2",
    "titre": "Route Lehri-Kerrouchen – 25 km",
    "fiche": [
        ("Marché", "n°46-RBK-2017 – Programme PRR3"),
        ("Nature", "Route rurale en zone montagneuse"),
        ("3 sections", "Linéaire (23 prix) + Carrefour (11 prix) + Bretelles (19 prix)"),
        ("53 prix", "au bordereau des prix"),
    ],
    "montant": "29 M DH TTC",
    "metres": [
        ("120 334 m³", "Déblais"),
        ("76 735 m³", "Remblais"),
        ("19 989 m³", "Couche base GNB"),
        ("34 481 m³", "Fondation GNF2"),
        ("794 ml", "Buses Ø1000"),
        ("789 m³", "Gabions"),
    ],
}

# =============================================================================
# PAGE 18 – BUDGET ROUTE (= page 24 actuelle + image CPS)
# =============================================================================
PAGE_18 = {
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
}

# =============================================================================
# PAGE 19 – SITUATION 5 : Cubatures de terrassement en montagne
# =============================================================================
SITUATION_5 = {
    "numero": 5,
    "titre": "Contrôle des cubatures de terrassement en zone montagneuse",
    "chiffre_cle": "120 334 m³",
    "chiffre_label": "Déblais vérifiés — surcoût absorbé sans avenant",
    "competence": "Vérifier les quantités exécutées et contrôler les écarts",
    "contexte": (
        "La route Lehri-Kerrouchen traverse 25 km de Moyen Atlas. Je "
        "devais vérifier les cubatures calculées par le bureau d'études : "
        "120 334 m³ de déblais et 76 735 m³ de remblais. Sur un projet "
        "routier de cette envergure, le poste terrassement est le plus "
        "sensible aux écarts — et celui où l'enveloppe budgétaire peut "
        "déraper le plus vite si le contrôle n'est pas rigoureux."
    ),
    "probleme": (
        "Au-delà du kilomètre 12, le terrain s'est révélé rocheux — du "
        "calcaire fracturé que l'étude géotechnique n'avait pas détecté. "
        "Résultat : 5 000 m³ à reclassifier en déblais rocheux, avec un "
        "surcoût immédiat de 285 000 DH. En parallèle, trois passages "
        "en fond de vallée nécessitaient des ouvrages hydrauliques non "
        "prévus au marché. L'écart cumulé montait à 8% du budget."
    ),
    "action": (
        "J'ai instauré un contrôle contradictoire tous les 500 mètres "
        "avec le conducteur de travaux : relevé GPS sur le terrain, "
        "comparaison avec les profils théoriques, calcul de volume "
        "vérifié par tableur. Pour les ouvrages hydrauliques, j'ai "
        "travaillé directement avec le bureau d'études pour dimensionner "
        "trois dalots béton armé. La reclassification des 5 000 m³ "
        "a été formalisée par attachement signé et documentée en photos."
    ),
    "resultat": (
        "Le surcoût de 285 000 DH a été entièrement absorbé par "
        "compensation sur d'autres postes — remblais réels inférieurs "
        "aux prévisions et économies sur le revêtement. Les trois "
        "dalots ont été intégrés dans l'enveloppe globale. La réception "
        "provisoire a eu lieu dans les délais contractuels, sans avenant "
        "et sans réclamation de l'entreprise. Ce contrôle m'a appris "
        "qu'il faut aller sur le terrain, systématiquement."
    ),
}

# =============================================================================
# PAGE 20 – DÉFIS ET SOLUTIONS P2 (= page 25 actuelle)
# =============================================================================
PAGE_20 = {
    "titre": "DÉFIS D'UN CHANTIER ROUTIER\nEN ZONE MONTAGNEUSE",
    "defis": [
        ("Relief montagneux",
         "Le tracé de 25 km traverse un terrain accidenté du Moyen Atlas avec un dénivelé cumulé de 400 m "
         "et des pentes atteignant 12% en lacets. Les volumes de terrassement étaient considérables : "
         "120 334 m³ de déblais et 76 735 m³ de remblais — au-delà des estimations initiales. "
         "Les ouvrages de soutènement en gabions (789 m³) s'avéraient indispensables pour stabiliser les talus. "
         "Toute erreur de calcul des cubatures se traduisait directement par un dépassement budgétaire.",
         "Contrôle méthodique par profils en travers tous les 25 m, complété par relevés GPS contradictoires "
         "tous les 500 ml avec le conducteur de travaux. Vérification de chaque volume par tableur avant validation. "
         "Ajustement du tracé sur 3 virages critiques en concertation avec le BET. "
         "La rigueur du contrôle a permis d'absorber les surplus de déblais grâce aux économies "
         "sur les remblais réels — inférieurs aux prévisions de l'étude."),
        ("Conditions climatiques",
         "La zone Moyen Atlas est soumise à un gel hivernal sévère (–8°C en janvier) "
         "et à des précipitations de 600 mm/an, concentrées entre novembre et mars. "
         "Les enrobés ne se posent pas sous le gel ; les terrassements en terrain détrempé créent des instabilités. "
         "Deux arrêts hivernaux formalisés par ordres de service, représentant 45 jours cumulés d'arrêt. "
         "Sans planification rigoureuse, le risque de dépassement du délai contractuel était réel.",
         "Planification rigoureuse calée sur la saison sèche (avril-octobre). Terrassements prioritaires "
         "en début de campagne, avant les premières pluies d'automne. Pose des enrobés réservée aux mois "
         "d'été (juin-août) pour des températures optimales. Protection des talus par géotextile avant l'hiver. "
         "OS d'arrêt et de reprise formalisés avec photos datées. Planning de rattrapage établi dès la reprise. "
         "Délai contractuel respecté malgré les 45 jours d'interruption cumulés."),
        ("Terrain rocheux imprévu",
         "L'étude géotechnique initiale ne signalait aucune zone rocheuse significative. "
         "À partir du PK 12+000, le terrassement a révélé du calcaire fracturé — non identifié. "
         "5 000 m³ reclassés en déblais rocheux (85 DH/m³ vs 28 DH/m³ pour les déblais ordinaires), "
         "soit un surcoût immédiat de 285 000 DH (+1,2%). L'écart cumulé avec les autres postes "
         "montait à +8% — bien au-delà du seuil d'avenant de 5%. "
         "Un avenant aurait bloqué le chantier pendant 3 à 6 mois de procédure administrative.",
         "Reclassification formalisée par attachement contradictoire signé des deux parties, "
         "documentée par photos du front de taille et mesures géoréférencées. "
         "Concertation avec le BET pour la révision des profils de terrassement impactés. "
         "Compensation intégrale par les économies réalisées sur les remblais réels "
         "(quantités inférieures aux prévisions) et sur une section de revêtement. "
         "Résultat : surcoût de 285 000 DH entièrement absorbé — réception sans avenant."),
        ("Ouvrages hydrauliques",
         "Trois talwegs (aux PK 15+500, 18+200 et 21+800) ont nécessité des ouvrages de franchissement "
         "non prévus au marché initial. Le dimensionnement des buses initialement prévu s'avérait insuffisant "
         "face aux débits observés pendant les travaux. Des dalots 3×2×2 m en béton armé étaient indispensables. "
         "Traités comme travaux supplémentaires, ils auraient représenté 28,1% du budget total — "
         "un avenant inévitable, engageant la responsabilité technique de l'Agence.",
         "Dimensionnement rigoureux avec le BET NOVEC sur la base du débit centennal Q100. "
         "Calcul des armatures HA12+HA16 validé par le bureau structure. "
         "Repositionnement stratégique des buses Ø1000 initialement prévues pour optimiser l'hydraulique. "
         "Intégration de l'ensemble dans l'enveloppe globale par compensation sur d'autres postes. "
         "Réception provisoire dans les délais contractuels — aucun avenant, aucune réclamation de l'entreprise."),
    ],
}

# =============================================================================
# PAGE 21 – ACTIVITÉS COMPLÉMENTAIRES (= page 26 actuelle)
# =============================================================================
PAGE_21 = {
    "titre": "ACTIVITÉS COMPLÉMENTAIRES",
    "maroc_titre": "5 AUTRES MARCHÉS AU CONSEIL RÉGIONAL",
    "marches": [
        ("27-RBK-2017", "Route Sidi Bouabbad → Oued Grou (12 km)", "Route rurale | 8,2 M DH"),
        ("28-RBK-2017", "Route Ajdir-Ayoun + Piste Lijon Kichchon", "Route+Piste | 6,5 M DH"),
        ("30-RBK-2017", "AEP El Borj → El Hamam (18 km de conduite)", "Eau potable | 4,8 M DH"),
        ("39-RBK-2017", "Pistes Hartaf → Sebt Ait Rahou (9 km)", "Pistes rurales | 3,1 M DH"),
        ("49-RBK-2016", "Aménagement voie Amghass → Bouchbel", "Voirie urbaine | 5,9 M DH"),
    ],
    "france_titre": "EXPÉRIENCE TERRAIN EN FRANCE (2022-2024)",
    "france": [
        ("2022-2023", "Chef d'équipe GO – Ergalis BTP, Feurs (Loire)",
         "Lecture plans, implantation/traçage, montage banches, armatures HA, coulage, cycles coffrage — GO structurel"),
        ("2024", "Chef de chantier – Minssieux et Fils, Mornant (Rhône)",
         "Encadrement opérationnel d'équipe, planning quotidien, avancement, contrôle qualité béton et maçonnerie"),
    ],
}

# =============================================================================
# PAGE 22 – SÉPARATEUR "Bilan et Analyse"
# =============================================================================
PAGE_22 = {
    "numero": "04",
    "titre": "BILAN ET ANALYSE",
    "sous_titre": "Synthèse des acquis professionnels | Comparaison Maroc/France\nBilan réflexif | Protocole BIM",
}

# =============================================================================
# PAGE 23 – TABLEAU SYNTHÈSE ACTIVITÉS → COMPÉTENCES (nouveau)
# =============================================================================
PAGE_23 = {
    "titre": "SYNTHÈSE DES ACTIVITÉS PROFESSIONNELLES",
    "tableau": [
        # (Activité, Savoir-faire mobilisé, Situation, Niveau)
        ("Métrés détaillés : 112 prix × 4 communes, surfaces/linéaires/volumes",
         "Réaliser des métrés tous corps d'état", "Situation 1", "Maîtrise"),
        ("Estimation confidentielle 15,8 M DH, croisement 3 sources de prix",
         "Estimer un ouvrage en phase AO", "Situation 1", "Maîtrise"),
        ("Analyse 3 offres, grille /100 (technique 60 + financier 40), PV CAO",
         "Analyser les offres en commission", "Situation 2", "Maîtrise"),
        ("Tableau de bord hebdomadaire 3 indicateurs, compensation inter-postes",
         "Suivre l'exécution financière", "Situation 3", "Expert"),
        ("Attachements contradictoires /500 ml, contrôle cubatures GPS",
         "Vérifier les quantités exécutées", "Situation 5", "Expert"),
        ("CR consolidé sous 24h, OS d'arrêt/reprise, notes factuelles",
         "Communiquer par écrit", "Situation 4", "Maîtrise"),
        ("Réunions chantier, point quotidien 8h, présentation CAO",
         "Communiquer oralement", "Situations 2, 4", "Maîtrise"),
        ("Rédaction CPS, RC, BPDE, vérification cohérence pièces/plans",
         "Rédiger pièces de marchés publics", "Situations 1-4", "Maîtrise"),
        ("Convention BIM LOD 300, export IFC 2x3, extraction 78 postes quantités",
         "Collaborer en BIM (Open BIM)", "Protocole BIM p.26", "Maîtrise"),
    ],
    "legende": "Niveau : Maîtrise = pratique régulière et autonome | Expert = transmission aux collaborateurs",
}

# =============================================================================
# PAGE 24 – COMPARAISON MAROC / FRANCE (= page 28 actuelle)
# =============================================================================
PAGE_24 = {
    "sur_titre": "ANALYSE COMPARATIVE",
    "titre": "MAROC vs FRANCE",
    "intro": (
        "Les deux systèmes partagent les mêmes principes fondamentaux : transparence, égalité de traitement "
        "des candidats, mise en concurrence et choix de l'offre économiquement la plus avantageuse. "
        "Les documents diffèrent (CPS/RC/BPDE au Maroc vs CCAP/CCTP/BPU-DQE en France), mais la logique "
        "de protection du maître d'ouvrage et de l'entreprise est identique. "
        "Cette double expérience développe une capacité d'adaptation précieuse dans un secteur BTP "
        "de plus en plus international — et change profondément la façon d'estimer, "
        "de rédiger les marchés et d'analyser les offres des entreprises."
    ),
    "comparaison": [
        ("Réglementation", "Décret n°2-12-349 (20/03/2013)", "Code de la commande publique (2019)"),
        ("Seuils", "AO ouvert > 500 000 DH", "AO au-dessus des seuils européens"),
        ("Pièces du marché", "CPS + RC + BPDE + Plans", "CCAP + CCTP + BPU/DQE ou DPGF"),
        ("Estimation", "Estimation confidentielle obligatoire", "Estimation du maître d'ouvrage"),
        ("Normes", "Normes marocaines, RPS 2000 (sismique)", "DTU, Eurocodes, RE2020"),
        ("Suivi financier", "Attachements contradictoires, décomptes", "Situations de travaux mensuelles"),
        ("Révision des prix", "Formule de révision contractuelle", "Actualisation et révision (CCAP)"),
        ("Commission", "Commission d'Appel d'Offres (CAO)", "Commission d'Appel d'Offres"),
    ],
    "synthese": "Mêmes fondamentaux de la commande publique — deux cadres réglementaires distincts — une seule exigence : la rigueur",
}

# =============================================================================
# PAGE 25 – BILAN RÉFLEXIF (nouveau)
# =============================================================================
PAGE_25 = {
    "titre": "BILAN RÉFLEXIF",
    "blocs": [
        ("Ce que j'ai appris",
         "La leçon la plus forte vient de la semaine 23, quand j'ai dû gérer trois crises simultanées "
         "sur quatre chantiers distants de 80 km. J'ai compris que l'écrit systématique — ordres de service, "
         "attachements contradictoires, comptes-rendus consolidés — est le seul rempart réel contre les litiges. "
         "Sans la trace écrite signée à Kerrouchen, le dépassement de 0,8% aurait pu être contesté par l'entreprise. "
         "J'ai aussi appris la valeur d'une estimation bien construite : croiser trois sources de prix m'a donné "
         "un écart de 3,2% là où les estimations de la région dépassaient souvent 10%. "
         "Et j'ai compris que comprendre les contraintes d'exécution depuis le terrain — coûts réels, "
         "rendements, aléas climatiques — rend les estimations infiniment plus justes."),
        ("Ce que je ferais différemment",
         "Avec le recul, j'aurais déployé le tableau de bord financier dès le premier mois — pas au neuvième, "
         "quand le signal d'alerte était déjà à +4,8%. Le problème détecté tôt coûte dix fois moins cher à corriger. "
         "J'aurais aussi insisté pour une étude géotechnique complémentaire avant les terrassements : "
         "la reclassification tardive de 5 000 m³ en déblais rocheux (285 000 DH de surcoût) "
         "aurait pu être anticipée et budgétée dès le lancement. "
         "Et j'aurais standardisé mes comptes-rendus dès le démarrage du chantier. "
         "Celui que j'ai créé dans l'urgence de la semaine 23 est devenu le modèle de l'Agence — "
         "preuve qu'un bon outil, même créé sous pression, finit toujours par s'imposer."),
        ("Ce que j'apporte au BTS MEC",
         "Mon expérience côté maîtrise d'ouvrage m'a appris que l'estimation confidentielle est l'acte "
         "fondateur de tout marché public — c'est elle qui conditionne la viabilité du projet. "
         "En France, j'ai retrouvé les mêmes fondamentaux avec le DQE et le DPGF, dans un cadre normatif différent. "
         "Peu de professionnels de l'économie de la construction combinent une expérience significative "
         "côté maîtrise d'ouvrage, côté exécution et une maîtrise du BIM. "
         "Le BTS MEC formalise ce que j'ai appris sur le terrain : la chaîne qui va du métré au décompte, "
         "en maîtrisant chaque maillon. Le BIM me permet aujourd'hui d'automatiser cette chaîne "
         "et de la rendre plus fiable — c'est sur cette triple conviction que j'ai fondé BIMCO."),
    ],
}

# =============================================================================
# PAGE 26 – PROTOCOLE BIM (nouveau)
# =============================================================================
PAGE_26 = {
    "titre": "PROTOCOLE DE COLLABORATION BIM",
    "sous_titre": "Appliquer un protocole de collaboration BIM",
    "convention": {
        "titre": "Convention BIM appliquée",
        "items": [
            ("Niveaux de détail", "LOD 300 (géométrie précise pour chiffrage) / LOI 3 (données matériaux, performances)"),
            ("Format d'échange", "IFC 2x3 – Open BIM | MVD Coordination View 2.0"),
            ("Plateforme", "Serveur collaboratif BIM360 avec gestion des droits et historique des versions"),
            ("Nomenclature", "Phase_Discipline_Lot_Niveau (ex : EXE_STR_LOT02_R+1)"),
        ],
    },
    "workflow": [
        "1. Modélisation Revit Architecture + Structure",
        "2. Export IFC et vérification de conformité",
        "3. Coordination et détection de clashs (Navisworks)",
        "4. Résolution des conflits entre disciplines",
        "5. Extraction automatique des quantités (Revit + Dynamo)",
        "6. Chiffrage et reporting avec traçabilité maquette",
    ],
    "cas_concret": {
        "titre": "CAS APPLIQUÉ : Bâtiment R+2 (Formation AFPA Colmar)",
        "details": (
            "Maquette Revit d'un bâtiment R+2 (logements collectifs, "
            "850 m² SHAB). Extraction automatique de 78 postes de métrés : "
            "surfaces de planchers, volumes béton (fondations, poteaux, "
            "poutres), linéaires de murs et quantités d'acier. Résultat : "
            "un écart de seulement 1,8% avec le métré manuel traditionnel. "
            "La détection de clashs en amont (12 conflits structure/réseaux, "
            "tous résolus) illustre la valeur ajoutée concrète du processus "
            "pour l'économiste."
        ),
    },
    "apport_mec": (
        "Le BIM transforme la chaîne métré → estimation → chiffrage. "
        "Sur le cas AFPA : extraction automatique en 2h vs 2 jours en métré "
        "traditionnel, écart de seulement 1,8%. Avantages pour l'économiste "
        "MEC : estimations plus fiables (données issues du modèle 3D), "
        "traçabilité totale (chaque quantité liée à un objet BIM), mise "
        "à jour instantanée (modification du modèle → recalcul automatique "
        "des quantités). Vision BIMCO : développer des plugins Revit/Dynamo "
        "pour automatiser le passage maquette → DPGF."
    ),
}

# =============================================================================
# PAGE 27 – PROJET PROFESSIONNEL (= page 29 actuelle)
# =============================================================================
PAGE_27 = {
    "titre": "MON PROJET PROFESSIONNEL",
    "horizons": [
        ("COURT TERME", "2026",
         "Obtenir le BTS MEC — validation officielle du parcours terrain\n"
         "Premières prestations BIMCO :\n"
         "· Métrés BIM et traditionnels\n"
         "· Études de prix, DPGF, estimations\n"
         "· Suivi financier de marchés\n"
         "App 'Gestion Chantiers' déjà livrée\n"
         "(React / Node.js / PostgreSQL)\n"
         "Premiers plugins Revit/Dynamo"),
        ("MOYEN TERME", "2027-28",
         "Gamme d'outils BIM pour le MEC :\n"
         "· Plugin Revit → DPGF automatisé\n"
         "· Chiffrage assisté par maquette\n"
         "· Base prix connectée (Batiprix)\n"
         "· App web suivi économique interactif\n"
         "· Génération automatique CCTP,\n"
         "  rapports CAO, bordereaux de prix\n"
         "3 à 5 clients récurrents visés"),
        ("LONG TERME", "2029+",
         "BIMCO = cabinet d'ingénierie\nBIM + Économie de la construction\n"
         "Axe 1 : Prestations\n"
         "· Métrés BIM, études de prix\n"
         "· AMO économique, OPC, analyse offres\n"
         "Axe 2 : Édition d'outils\n"
         "· SaaS pour économistes MEC\n"
         "· Plugins Revit distribués\n"
         "Équipe 3-5 — CA 200-300 k€/an"),
    ],
    "citation": "« Les outils numériques doivent être\nau service de l'économiste de la construction,\net non l'inverse. »",
}

# =============================================================================
# PAGE 28 – CONCLUSION (mise à jour)
# =============================================================================
PAGE_28 = {
    "titre": "CONCLUSION",
    "resume": (
        "Ce rapport est avant tout une démonstration. Non pas de ce que je sais, "
        "mais de comment je résous un problème sous contrainte — délai, budget, "
        "aléas géologiques, coordination de crise. À chaque situation, j'ai dû "
        "trouver la bonne réponse avec les moyens disponibles. Ce que j'ai "
        "construit en huit ans, c'est une méthode. Le BTS MEC lui donne un cadre "
        "formel. BIMCO lui donne une continuation."
    ),
    "kpis": [
        ("3,2%", "Écart estimation"),
        ("94/100", "Note en CAO"),
        ("+0,8%", "Dépassement final"),
        ("48 h", "3 crises résolues"),
        ("0 avt", "Avenant cubatures"),
    ],
    "points": [
        ("Cinq compétences construites sur le terrain",
         "Estimer, analyser les offres, suivre financièrement, coordonner en crise, "
         "contrôler les quantités — cinq réflexes professionnels acquis par la pratique, "
         "pas par la théorie"),
        ("Une double lecture des projets",
         "Côté maîtrise d'ouvrage au Maroc, côté exécution en France — comprendre "
         "les enjeux des deux côtés de la table change profondément la façon "
         "d'estimer et de négocier"),
        ("La rigueur de l'écrit comme protection",
         "Ordres de service, attachements contradictoires, comptes-rendus consolidés : "
         "la trace écrite est le seul rempart réel contre les litiges en cours de chantier"),
        ("BIMCO : la traduction de ce parcours",
         "Un bureau où la rigueur du terrain rencontre les outils numériques — "
         "parce que le chiffrage mérite d'être aussi précis que la construction"),
    ],
    "citation": (
        "« Le BTS MEC ne valide pas seulement un diplôme.\n"
        "Il valide huit ans de terrain, de chiffres et de chantier.\n"
        "Et il ouvre la voie à ce que je veux construire :\n"
        "des outils qui changent le quotidien de l'économiste. »"
    ),
    "pied": "BAHAFID Mohamed | BIMCO | BTS MEC Session 2026 | Académie de Lyon",
}

# =============================================================================
# PAGE 29 – ANNEXE 1 : DOCUMENTS OFFICIELS (nouveau)
# =============================================================================
PAGE_29 = {
    "titre": "ANNEXE 1 : DOCUMENTS OFFICIELS",
    "documents": [
        ("cao", "Convocation CAO", "Commission d'Appel d'Offres\nRégion Béni Mellal-Khénifra"),
        ("rejet", "Notification de rejet", "Procédure marchés publics\nDécret n°2-12-349"),
        ("cps", "CPS signé", "Cahier des Prescriptions Spéciales\nMarché n°46-RBK-2017"),
    ],
}

# =============================================================================
# PAGE 30 – ANNEXE 2 : PHOTOS CHANTIER (nouveau)
# =============================================================================
PAGE_30 = {
    "titre": "ANNEXE 2 : PHOTOS DE CHANTIER",
    "photos": [
        ("chantier1", "Aménagement urbain – VRD"),
        ("chantier2", "Pose de bordures"),
        ("chantier3", "Terrassement"),
        ("chantier4", "Assainissement"),
        ("chantier5", "Enrobés"),
        ("chantier6", "Route montagne"),
        ("route", "Route Lehri-Kerrouchen"),
        ("route2", "Section montagneuse"),
        ("terrassement", "Terrassement rocheux"),
    ],
}

# =============================================================================
# TABLEAUX DE DONNÉES (issus du rapport DOCX de référence)
# =============================================================================

TABLE_CORPS_ETAT = {
    "titre": "8 corps d'état — Projet mise à niveau 4 communes",
    "colonnes": ["Partie", "Désignation", "Exemple de travaux"],
    "lignes": [
        ("01", "Assainissement & réseaux", "Tranchées, buses PEHD Ø400-Ø500, regards, bouches d'égout"),
        ("02", "Travaux de chaussée", "Terrassement, couche GNF1, couche base GNA, enrobés EB 0/10"),
        ("03", "Aménagement trottoirs", "Bordures T1/T3, béton, revêtement carreaux striés, pavés"),
        ("04", "Signalisation", "Marquage au sol, panneaux réglementaires, peinture bordures"),
        ("05", "Éclairage public", "Tranchées, tubes annelés, massifs candélabres (163 U), câbles"),
        ("06", "Murs & ouvrages divers", "Béton armé, maçonnerie moellons, murs de soutènement, gabions"),
        ("07", "Aménagement paysager", "Terre végétale, réseau d'arrosage intégré, plantation arbres"),
        ("08", "Mobilier urbain", "Corbeilles, bancs en granite, bornes d'aménagement"),
    ],
}

TABLE_BUDGET_COMMUNES = {
    "titre": "Budget par commune — Mise à niveau 4 communes",
    "colonnes": ["Commune", "Montant HT (DH)", "Montant TTC (DH)", "% du total"],
    "lignes": [
        ("El Hammam", "6 634 919", "7 961 903", "14,9%"),
        ("Kerrouchen", "7 336 914", "8 804 297", "16,5%"),
        ("Ouaoumana", "15 830 359", "18 996 431", "35,5%"),
        ("Sebt Ait Rahou", "14 803 389", "17 764 067", "33,2%"),
        ("TOTAL", "44 605 581", "53 526 697", "100%"),
    ],
    "note": "Source : marchés n°38 à 41-RBK-2017, Conseil Régional BMK",
}

TABLE_METRES_KERROUCHEN = {
    "titre": "Métrés principaux — Commune de Kerrouchen (extrait BPDE)",
    "colonnes": ["Ouvrage", "Quantité", "Unité"],
    "lignes": [
        ("Tranchées assainissement", "3 620", "m³"),
        ("Buses PEHD Ø400", "1 160", "ml"),
        ("Buses PEHD Ø500", "1 210", "ml"),
        ("Regards de visite", "56", "U"),
        ("Terrassement chaussée", "4 200", "m³"),
        ("Couche de fondation GNF1", "1 560", "m³"),
        ("Enrobés EB 0/10", "2 175", "T"),
        ("GBB 0/14", "1 245", "T"),
        ("Bordures T3", "4 800", "ml"),
        ("Bordures T1", "5 890", "ml"),
        ("Revêtement carreaux striés", "11 500", "m²"),
        ("Massifs candélabres", "163", "U"),
    ],
    "note": "Avant-métré établi en phase études — Commune de Kerrouchen (marché n°39-RBK-2017)",
}

TABLE_BUDGET_ROUTE = {
    "titre": "Budget par poste — Route Lehri-Kerrouchen (24,2 M DH HT)",
    "colonnes": ["Poste", "Montant HT (DH)", "%"],
    "lignes": [
        ("Terrassement (déblais + remblais)", "4 325 057", "17,9%"),
        ("Corps de chaussée (GNB + GNF2)", "7 280 884", "30,1%"),
        ("Revêtement (bicouche + imprégnation)", "2 588 570", "10,7%"),
        ("Ouvrages hydrauliques (buses + béton)", "6 793 920", "28,1%"),
        ("Ouvrages de soutènement (gabions)", "458 220", "1,9%"),
        ("Bretelles et carrefour", "2 722 720", "11,3%"),
        ("TOTAL HT", "24 169 371", "100%"),
    ],
}

TABLE_AUTRES_MARCHES = {
    "titre": "Autres marchés publics pilotés au Conseil Régional",
    "colonnes": ["N° Marché", "Objet", "Type / Montant"],
    "lignes": [
        ("27-RBK-2017", "Route village Sidi Bouabbad → Oued Grou (12 km)", "Route rurale | 8,2 M DH"),
        ("28-RBK-2017", "Route Ajdir-Ayoun + Piste Lijon Kichchon", "Route + Piste | 6,5 M DH"),
        ("30-RBK-2017", "Adduction Eau Potable El Borj – El Hamam (18 km)", "AEP | 4,8 M DH"),
        ("39-RBK-2017", "Pistes Hartaf – Sebt Ait Rahou (9 km)", "Pistes rurales | 3,1 M DH"),
        ("49-RBK-2016", "Aménagement voie Amghass – Bouchbel", "Voirie urbaine | 5,9 M DH"),
    ],
    "note": "Au total : 7 marchés publics pilotés pour + de 100 M DH d'investissements",
}

TABLE_COMPARAISON_REG = {
    "titre": "Comparaison réglementaire Maroc / France — 8 aspects",
    "colonnes": ["Aspect", "Maroc", "France"],
    "lignes": [
        ("Réglementation", "Décret n°2-12-349 (20/03/2013)", "Code de la commande publique"),
        ("Seuils procédure", "AO ouvert > 500 000 DH", "AO > seuils européens"),
        ("Pièces marché", "CPS + RC + BPDE", "CCAP + CCTP + BPU/DQE ou DPGF"),
        ("Estimation", "Estimation confidentielle obligatoire", "Estimation du maître d'ouvrage"),
        ("Normes construction", "Normes marocaines, RPS 2000 (sismique)", "DTU, Eurocodes, RE2020"),
        ("Suivi financier", "Attachements contradictoires, décomptes", "Situations mensuelles"),
        ("Révision des prix", "Formule de révision contractuelle", "Actualisation et révision (CCAP)"),
        ("Commission", "Commission d'Appel d'Offres (CAO)", "Commission d'Appel d'Offres"),
    ],
}

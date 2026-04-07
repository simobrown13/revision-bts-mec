from fpdf import FPDF

# Définition des couleurs de la charte graphique
BLEU_MARINE = (30, 50, 90)
ORANGE = (242, 145, 0)
BLEU_CLAIR = (120, 190, 210)
GRIS = (100, 100, 100)

class RapportPDF(FPDF):
    def header(self):
        # Ne pas mettre d'en-tête sur la page de garde
        if self.page_no() > 1:
            self.set_fill_color(*BLEU_MARINE)
            self.rect(0, 0, 210, 15, 'F')
            self.set_xy(10, 4)
            self.set_text_color(255, 255, 255)
            self.set_font('Arial', 'B', 10)
            self.cell(0, 8, 'BTS MEC - SESSION 2026', 0, 0, 'L')

    def footer(self):
        # Bandeau de pied de page
        self.set_y(-20)
        self.set_fill_color(*BLEU_MARINE)
        self.rect(0, 280, 210, 17, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', '', 9)
        self.cell(0, 10, f'BAHAFID Mohamed | Rapport U62 | Page {self.page_no()}/30', 0, 0, 'R')

    def titre_page(self, titre, couleur=BLEU_MARINE):
        self.set_xy(15, 30)
        self.set_text_color(*couleur)
        self.set_font('Arial', 'B', 22)
        self.multi_cell(0, 10, titre, align='L')
        self.ln(10)

    def texte_courant(self, texte):
        self.set_x(15)
        self.set_text_color(50, 50, 50)
        self.set_font('Arial', '', 12)
        self.multi_cell(180, 7, texte)
        self.ln(5)

# Initialisation du PDF
pdf = RapportPDF('P', 'mm', 'A4')
pdf.set_auto_page_break(auto=True, margin=25)

# --- PAGE 1 : PAGE DE GARDE ---
pdf.add_page()
pdf.set_fill_color(*BLEU_CLAIR)
pdf.circle(170, 50, 60, 'F') # Décoration graphique

pdf.set_xy(20, 80)
pdf.set_text_color(*BLEU_MARINE)
pdf.set_font('Arial', 'B', 40)
pdf.multi_cell(0, 15, "RAPPORT\nD'ACTIVITÉS\nPROFESSIONNELLES")

pdf.set_xy(20, 140)
pdf.set_text_color(*ORANGE)
pdf.set_font('Arial', 'B', 16)
pdf.cell(0, 10, "CANDIDAT", ln=True)

pdf.set_text_color(*BLEU_MARINE)
pdf.set_font('Arial', 'B', 30)
pdf.cell(0, 12, "BAHAFID", ln=True)
pdf.set_text_color(*ORANGE)
pdf.cell(0, 12, "Mohamed", ln=True)

pdf.set_xy(100, 145)
pdf.set_font('Arial', '', 12)
pdf.set_text_color(*GRIS)
pdf.multi_cell(90, 6, "BTS MEC\nManagement Economique\nde la Construction", align='R')


# --- PAGE 2 : IDENTITÉ ---
pdf.add_page()
pdf.titre_page("FICHE D'IDENTITE DU CANDIDAT")
infos = (
    "Candidat : BAHAFID Mohamed\n"
    "N° Candidat : 02537399911\n"
    "Académie : Lyon\n\n"
    "Structure d'accueil : Conseil Régional de Béni Mellal-Khénifra (Maroc)\n"
    "Direction : Agence d'Exécution des Projets\n"
    "Poste occupé : Technicien Études et Suivi des Travaux\n"
    "Durée d'expérience : 8 ans dans le BTP (3 ans Maroc + 5 ans France)\n\n"
    "Formation BIM : Technicien Modeleur BIM - AFPA Colmar\n"
    "Activité actuelle : BIMCO - Projeteur BIM / Économiste"
)
pdf.texte_courant(infos)


# --- PAGE 3 : SOMMAIRE ---
pdf.add_page()
pdf.titre_page("SOMMAIRE")
sommaire = (
    "1. Introduction et parcours (p. 4)\n\n"
    "2. Cadre professionnel (p. 5-7)\n\n"
    "3. Projet 1 : Mise à niveau 4 communes (p. 8-15)\n\n"
    "4. Projet 2 : Route Lehri-Kerrouchen (p. 16-20)\n\n"
    "5. Activités complémentaires (p. 21)\n\n"
    "6. Bilan et analyse (p. 22-26)\n\n"
    "7. Projet professionnel (p. 27)\n\n"
    "8. Conclusion et annexes (p. 28-30)"
)
pdf.texte_courant(sommaire)


# --- GENERATION DES PAGES RESTANTES (4 à 30) ---
# Pour l'exemple, nous générons une structure vide pour les pages suivantes
for i in range(4, 31):
    pdf.add_page()
    if i == 5:
        pdf.titre_page("01 - CADRE PROFESSIONNEL", ORANGE)
    elif i == 8:
        pdf.titre_page("02 - PROJET 1 : MISE À NIVEAU DE 4 COMMUNES", ORANGE)
    elif i == 11:
        pdf.titre_page("SITUATION 1 : Estimation confidentielle")
        pdf.texte_courant("CONTEXTE :\nÉtablir l'estimation de l'administration pour le marché d'Ouaoumana (15,8 M DH).\n\nACTION :\nMétrés croisés, actualisation de la mercuriale.\n\nRESULTAT :\nÉcart de 3,2% avec l'offre retenue.")
    elif i == 16:
        pdf.titre_page("03 - PROJET 2 : ROUTE LEHRI-KERROUCHEN", ORANGE)
    elif i == 22:
        pdf.titre_page("04 - BILAN ET ANALYSE", ORANGE)
    elif i == 29:
        pdf.titre_page("ANNEXE 1 : DOCUMENTS OFFICIELS")
        pdf.texte_courant("[Insérez ici vos documents officiels lors de l'édition du PDF]")
    elif i == 30:
        pdf.titre_page("ANNEXE 2 : PHOTOS DE CHANTIER")
        pdf.texte_courant("[Insérez ici vos photos de chantiers]")
    else:
        pdf.titre_page(f"Contenu de la page {i}")
        pdf.texte_courant("Emplacement pour le contenu détaillé du rapport (tableaux, analyses, schémas CPAR).")

# Génération du fichier
nom_fichier = "Rapport_BAHAFID_Mohamed.pdf"
pdf.output(nom_fichier)
print(f"Le fichier {nom_fichier} a été généré avec succès avec 30 pages !")
# FICHE MÉMO BTS MEC — À IMPRIMER RECTO-VERSO (A4)

---

# ═══════════ RECTO — ÉPREUVE E4 ═══════════

## THERMIQUE (Fiche 04)

| Formule | Usage |
|---|---|
| **R = e / λ** | Résistance d'une couche [m².K/W] — e en MÈTRES |
| **RT = Rsi + R1+...+Rn + Rse** | Résistance totale de la paroi |
| **U = 1 / RT** | Coefficient de transmission [W/m².K] |
| **a = λ / (ρ × Cp)** | Diffusivité thermique [m²/s] |
| **φ = 0,023 × e / √a** | Déphasage [heures] → doit être **> 12h** |

**Rsi/Rse :** Paroi verticale → Rsi = 0,13 / Rse = 0,04 | Plancher haut (flux ↓) → Rsi = 0,17

**RE2020 :** U mur ≤ 0,20 W/m².K | U toiture ≤ 0,15 W/m².K | Bbio + Cep + Ic énergie + Ic construction

**Matériaux clés :**
```
Laine de verre : λ = 0,032-0,040 | Paille de riz : λ = 0,056 / ρ = 120 / Cp = 1600
Laine de bois  : λ = 0,038-0,052 | Béton         : λ = 1,75 / ρ = 2400
PSE            : λ = 0,029-0,038 | Bois massif   : λ = 0,12-0,17
```
**Exemple paille de riz (e=0,37m) :** a = 2,92×10⁻⁷ → φ = 15,8h ✔

---

## ACOUSTIQUE (Fiche 05)

| Catégorie voie | Niveau | Distance | DnT,A,tr min (logement) |
|---|---|---|---|
| 1 | > 81 dB(A) | 300 m | 45 dB |
| 2 | 76-81 dB(A) | 250 m | 42 dB |
| **3** | **70-76 dB(A)** | **100 m** | **38 dB** |
| **4** | **65-70 dB(A)** | **30 m** | **35 dB** |
| 5 | 60-65 dB(A) | 10 m | 30 dB |

**Entrées d'air :** Standard AVENT = 30 dB | Acoustique AVENT = 35-44 dB | SMEC = > 44 dB

**Temps de réverbération (Sabine) :**
```
TR = 0,16 × V / A    A = Σ(Si × αi)    A_nec = 0,16 × V / TR_cible
```
**TR cibles :** Classe = 0,4-0,6 s | Gymnase = 1,5-2,0 s | Restaurant = 0,8-1,2 s

**α clés :** Laine minérale plafond = 0,70-0,90 | Moquette = 0,25-0,40 | Béton brut = 0,02

---

## FONDATIONS (Fiche 03)

**Hors gel :** H1 (montagne) = 1,00 m | H2 (plaine) = **0,80 m** | H3 (mer) = 0,50 m

| Classe | Désignation | Exemple fréquent |
|---|---|---|
| X0 | Aucun risque | Béton non armé intérieur sec |
| **XC2** | Humide, rarement sec | **Fondations, semelles enterrées** |
| **XC4** | Alternance humide/sec | **Façades exposées pluie** |
| **XD3** | Chlorures hors mer | **Parking (sels déverglaçage)** |
| **XS1** | Air marin (non immergé) | Côte < 1 km |
| **XF1** | Gel modéré | Éléments extérieurs, zone H2 |
| XA1 | Chimique faible | Sol peu agressif |

**Exigences béton :** XC2 → C25/30, w/c ≤ 0,60, 280 kg/m³, enrobage 25mm | XC4 → C30/37, w/c ≤ 0,50, 300 kg/m³, 30mm

**⚠ Un élément peut cumuler plusieurs classes : ex. poteau extérieur = XC4 + XF1**

---

## MOE / RÉGLEMENTATION / SÉCURITÉ (Fiches 01-02-06)

**Phases MOE :** ESQ → APS → APD → PRO → DCE → ACT → DET → AOR (+OPC)

**MOE composition :** Architecte = Mandataire | BET Structure/Fluides/QEB = Co-traitants

**Barème VT :** Non traité=0 | Peu détaillé=3 | Détaillé=6 | Bien=8 | Excellent=10

**PC :** Validité 3 ans + prorogation 2×1an | RT2012 si PC avant 01/01/2022 | RE2020 si après

**ERP :** R=Enseignement | L=Spectacles | M=Magasins | N=Restaurants | P=Sport | U=Santé | W=Admin

**MH :** Zone protection = **500 m** | Accord ABF obligatoire si visible

**Blindage :** Obligatoire si profondeur ≥ **1,30 m** | Matériaux = min **0,60 m** du bord | Engins lourds = **3 m**

**Garde-corps :** H ≥ 1,00 m (bord accessible) | Fixations : sabots / contreplaque+tirant / lestage

---

# ═══════════ VERSO — ÉPREUVE E52 ═══════════

## RÉVISION DES PRIX (Fiche 08)

```
Cn = 0,15 + 0,85 × (Im / I0)           ← arrondi au MILLIÈME SUPÉRIEUR
Révision HT = Travaux HT × (Cn − 1)
Formule 2 indices : Cn = 0,15 + 0,85 × [(Im1/I0_1)×p1 + (Im2/I0_2)×p2]
Actualisation simple : P = P0 × (Im-3 / I0)
```

**Im-3** = indice du mois M-3 (3 mois AVANT le mois des travaux)

**Indices BT clés :** BT01=tous CE | BT06=béton armé | BT30=ardoise | BT34=zinc | BT42=menuiserie acier

**⚠ Jamais de révision sur l'avance forfaitaire ! | Im < I0 → révision NÉGATIVE (déduction)**

---

## AVANCE FORFAITAIRE et RETENUE DE GARANTIE (Fiche 09)

| Règle | Formule |
|---|---|
| **Conditions avance** | Marché > 50 000 € HT **ET** durée > 2 mois |
| **Avance ≤ 12 mois** | 10% × Montant initial **TTC** |
| **Avance > 12 mois** | 10% × (12 × TTC / N mois) |
| **Remb. cumulé** | Avance × **(% cum − 65%) / 35%** |
| **Remb. mensuel** | R_cumulé − Remboursements déjà effectués |
| **Début remb.** | Quand % cumulé **≥ 65%** |
| **Fin remb.** | Quand % cumulé = **100%** |
| **RG mensuelle** | 5% × Montant **TTC** du décompte |
| **Plafond RG** | 5% × Montant initial **TTC** du marché |

---

## DÉCOMPTE MENSUEL — 16 LIGNES (Fiche 09)

```
1.  Avance forfaitaire
2.  Approvisionnements
3.  Travaux HT du mois  [= Marché HT × (% cum M − % cum M-1)]
4.  TOTAL créditeur (1+2+3)
5.  Remboursement avance  ← AVANT TVA
6.  Pénalités (1/3000 × HT × jours)
7.  Total débiteur (5+6)
8.  Solde HT (4−7)
9.  Révision des prix [Travaux × (Cn−1)]
10. Total HT (8+9)
11. TVA 20%
12. Total TTC (10+11)
13. Retenue de garantie [5% × TTC]  ← SUR LE TTC
14. Solde TTC net (12−13)
15. Acomptes antérieurs
16. ACOMPTE DU MOIS (14−15)
```

---

## ANALYSE DES OFFRES (Fiche 07)

```
Montant final  = Offre − Moins-values + Plus-values − Remise
Note prix      = N_max × (Prix_min / Prix_i)
Note délai     = N_max × (Délai_min / Délai_i)
Écart %        = (Offre − Estimation) / Estimation × 100
Ratio €/m²    = Montant HT lot / Surface (m²)
Score total    = Σ toutes les notes pondérées → offre avec score MAX retenue
```

**Pondérations rencontrées :** Prix 30-60% | VT 30-70% | Délai 0-20%

**VT sous-critères :** Note pondérée = Note/10 × Poids | Barème : Non traité=0 / Peu=3 / Bien=8 / Excellent=10

---

## RENTABILITÉ (Fiche 10)

```
DSMO = Heures × Taux horaire chargé (€/h)
DS   = DSMO + DSMAT
CD   = DS + Sous-traitance + Frais chantier directs
MB   = Montant HT − CD          MB% sain = 15-25%
MBH  = MB / Heures              MBH saine = 15-30 €/h (GO)
FOp  = ~0,90% × Montant HT
RB   = MB − FG − FOp            RB% sain = 2-8%
RB%  = RB / Montant HT × 100
```

**⚠ RB% < 2% = risque | RB% > 10% = excellente rentabilité**

---

## RAPPELS CRITIQUES — NE PAS OUBLIER

```
E4 :
□ e toujours en MÈTRES (140 mm = 0,140 m)
□ Hors gel H2 = 0,80 m (pas 1 m)
□ Cumuler les classes béton : XC4 + XF1 possible
□ Entrée d'air = maillon faible de l'isolement façade
□ PC validité 3 ans + 2×1 an prorogation

E52 :
□ Cn arrondi au MILLIÈME SUPÉRIEUR (1,03294 → 1,033)
□ Im-3 = indice 3 mois AVANT le mois des travaux
□ Avance = 10% du TTC (pas du HT)
□ Remb. avance commence à 65% (pas avant !)
□ Remb. avance déduit AVANT TVA — RG calculée SUR le TTC
□ Note prix utilise Prix_min APRÈS négociation
```

# EXERCICES CORRIGÉS — RENTABILITÉ D'OPÉRATION (E52)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C17-3 | Savoir S7

---

## RAPPEL DES FORMULES

```
DÉBOURSÉ SEC (DS)
  DSMO  = Heures productives × Taux horaire chargé (€/h)
  DSMAT = Coût des matériaux mis en œuvre (€ HT)
  DS    = DSMO + DSMAT

COÛT DIRECT (CD)
  CD = DS + Sous-traitance + Frais de chantier directs

MARGE BRUTE (MB)
  MB  = Montant HT travaux - CD
  MBH = MB / Heures productives    [€/h]  = marge brute horaire

RÉSULTAT BRUT (RB)
  FOp = % × Montant HT    (Frais d'Opération, ~1%)
  FG  = % × Montant HT    (Frais Généraux de l'entreprise)
  RB  = MB - FG - FOp
  RB% = RB / Montant HT × 100

RATIOS ÉCONOMIQUES
  Ratio (€/m²)    = Montant HT lot / Surface de référence
  Ratio (€/unité) = Montant HT / Nombre d'unités
```

---

## EXERCICE 1 — Calcul du Déboursé Sec (type sujet 2023 — École maternelle)

### Énoncé
Chantier de gros œuvre d'une école maternelle — lot béton armé.

**Main d'œuvre :**
| Ressource | Heures prod. | Taux horaire chargé |
|---|---|---|
| Chef de chantier | 120 h | 38,50 €/h |
| Maçon qualifié (×2) | 480 h | 32,00 €/h |
| Manœuvre (×2) | 360 h | 26,50 €/h |

**Matériaux :**
| Matériau | Quantité | Prix unitaire HT |
|---|---|---|
| Béton C25/30 | 85 m³ | 142 €/m³ |
| Aciers HA | 4 200 kg | 1,15 €/kg |
| Coffrages bois | 320 m² | 18,50 €/m² |
| Armatures treillis | 210 m² | 8,20 €/m² |

**Questions :**
1. Calculer DSMO (déboursé sec main d'œuvre).
2. Calculer DSMAT (déboursé sec matériaux).
3. Calculer DS total.

---

### CORRIGÉ

**1. DSMO :**

```
Chef chantier : 120 × 38,50  =  4 620,00 €
Maçons        : 480 × 32,00  = 15 360,00 €
Manœuvres     : 360 × 26,50  =  9 540,00 €

DSMO = 4 620 + 15 360 + 9 540 = 29 520,00 €

Total heures productives = 120 + 480 + 360 = 960 heures
```

**2. DSMAT :**

```
Béton       : 85     × 142,00 = 12 070,00 €
Aciers HA   : 4 200  ×   1,15 =  4 830,00 €
Coffrages   : 320    ×  18,50 =  5 920,00 €
Treillis    : 210    ×   8,20 =  1 722,00 €

DSMAT = 12 070 + 4 830 + 5 920 + 1 722 = 24 542,00 €
```

**3. DS total :**

```
DS = DSMO + DSMAT
DS = 29 520 + 24 542
DS = 54 062,00 € HT
```

---

## EXERCICE 2 — Marge brute et résultat (type sujet 2023)

### Énoncé
Suite de l'exercice 1. Données complémentaires :
- Montant HT du lot : **74 800 € HT**
- Sous-traitance (échafaudages) : **3 200 € HT**
- Frais de chantier directs (installation, outillage) : **2 100 € HT**
- Frais Généraux (FG) : **7% du montant HT**
- Frais d'Opération (FOp) : **1% du montant HT**

**Questions :**
1. Calculer le Coût Direct CD.
2. Calculer la Marge Brute MB et la Marge Brute Horaire MBH.
3. Calculer le Résultat Brut RB et RB%.
4. L'opération est-elle rentable ? (seuil de rentabilité acceptable : RB% ≥ 3%)

---

### CORRIGÉ

**1. Coût Direct :**

```
CD = DS + Sous-traitance + Frais chantier directs
CD = 54 062 + 3 200 + 2 100
CD = 59 362,00 € HT
```

**2. Marge Brute :**

```
MB = Montant HT - CD
MB = 74 800 - 59 362
MB = 15 438,00 € HT

MBH = MB / Heures productives
MBH = 15 438 / 960
MBH = 16,08 €/h
```

**3. Résultat Brut :**

```
FG  = 7% × 74 800 = 5 236,00 €
FOp = 1% × 74 800 =   748,00 €

RB  = MB - FG - FOp
RB  = 15 438 - 5 236 - 748
RB  = 9 454,00 € HT

RB% = RB / Montant HT × 100
RB% = 9 454 / 74 800 × 100
RB% = 12,64%
```

**4. Rentabilité :**

```
RB% = 12,64% > 3%  →  OPÉRATION RENTABLE
→ Très bon résultat pour un lot de gros œuvre.
→ La marge brute horaire de 16,08 €/h est satisfaisante.
```

---

## EXERCICE 3 — Analyse de rentabilité complète (type sujet 2025 — Halle Maubeuge)

### Énoncé
Lot charpente métallique d'une halle couverte 901 m².
- Montant HT marché : **142 500 € HT**
- Surface : 901 m²

**Main d'œuvre :**
| Équipe | Heures | Taux |
|---|---|---|
| Monteur charpente N3 (×3) | 720 h | 35,50 €/h |
| Grutier | 80 h | 41,00 €/h |
| Chef de chantier | 60 h | 42,00 €/h |

**Matériaux :**
| Poste | Montant HT |
|---|---|
| Charpente acier (profilés, boulons) | 48 200 € |
| Bardage et couverture | 12 800 € |
| Boulonnerie et consommables | 1 450 € |

**Autres :**
- Location grue : **6 800 € HT** (sous-traitance)
- Installation chantier : **1 200 € HT**
- FG : 8% — FOp : 1,5%

**Questions :**
1. Calculer DSMO, DSMAT, DS, CD.
2. Calculer MB, MBH, RB, RB%.
3. Calculer le ratio €/m² pour ce lot.

---

### CORRIGÉ

**1. DSMO, DSMAT, DS, CD :**

```
DSMO :
  Monteurs   : 720 × 35,50 = 25 560 €
  Grutier    :  80 × 41,00 =  3 280 €
  Chef ch.   :  60 × 42,00 =  2 520 €
  DSMO = 25 560 + 3 280 + 2 520 = 31 360 €
  Total heures = 720 + 80 + 60 = 860 h

DSMAT :
  Charpente acier = 48 200 €
  Bardage/couv.   = 12 800 €
  Boulonnerie     =  1 450 €
  DSMAT = 62 450 €

DS = 31 360 + 62 450 = 93 810 €

CD :
  CD = DS + Sous-traitance + Frais chantier
  CD = 93 810 + 6 800 + 1 200
  CD = 101 810 €
```

**2. MB, MBH, RB, RB% :**

```
MB  = 142 500 - 101 810 = 40 690 €
MBH = 40 690 / 860 = 47,31 €/h

FG  = 8,0% × 142 500 = 11 400 €
FOp = 1,5% × 142 500 =  2 137,50 €

RB  = 40 690 - 11 400 - 2 137,50 = 27 152,50 €
RB% = 27 152,50 / 142 500 × 100 = 19,05%
```

**3. Ratio €/m² :**

```
Ratio = Montant HT / Surface
Ratio = 142 500 / 901
Ratio = 158,16 €/m²
```

---

## EXERCICE 4 — Ratios économiques multi-lots (type sujet 2024)

### Énoncé
Gymnase municipal — 2 056 m² SHON. Récapitulatif des lots :

| Lot | Désignation | Montant HT |
|---|---|---|
| 1 | Gros œuvre | 612 000 € |
| 2 | Charpente | 287 500 € |
| 3 | Couverture | 143 200 € |
| 4 | Menuiseries ext. | 89 400 € |
| 5 | Plâtrerie-isolation | 76 800 € |
| 6 | Électricité | 112 300 € |
| 7 | Plomberie-CVC | 134 600 € |

**Questions :**
1. Calculer le montant total HT de l'opération.
2. Calculer le ratio global €/m² SHON.
3. Calculer le ratio du lot gros-œuvre en €/m².
4. Le ratio gros-œuvre est-il cohérent avec un bâtiment sportif (référence : 280 à 380 €/m²) ?

---

### CORRIGÉ

**1. Montant total :**

```
Total = 612 000 + 287 500 + 143 200 + 89 400 + 76 800 + 112 300 + 134 600
Total = 1 455 800 € HT
```

**2. Ratio global :**

```
Ratio global = 1 455 800 / 2 056
Ratio global = 708,08 €/m² SHON
```

**3. Ratio gros-œuvre :**

```
Ratio GO = 612 000 / 2 056
Ratio GO = 297,66 €/m²
```

**4. Analyse :**

```
Ratio GO calculé = 297,66 €/m²
Référence bâtiment sportif = 280 à 380 €/m²

297,66 est dans la fourchette [280 ; 380]
→ COHÉRENT avec un gymnase municipal standard.
```

---

## POINTS DE VIGILANCE

- **DSMO** = heures × taux **chargé** (charges sociales incluses dans le taux)
- **MBH** = indicateur clé de productivité (plus elle est élevée, plus c'est rentable)
- **FG et FOp** s'appliquent sur le **montant HT** (pas sur DS ou CD)
- Le ratio **€/m²** compare toujours au marché de référence pour évaluer la cohérence
- **RB% minimal acceptable** en bâtiment : environ 2 à 5% selon les entreprises

# EXERCICES CORRIGÉS — SITUATION DE TRAVAUX ET DÉCOMPTE MENSUEL (E52)
> Basés sur les sujets BTS MEC 2023-2025 | Compétence C17-2 | Savoir S7

---

## RAPPEL DES FORMULES

```
AVANCE FORFAITAIRE
  Conditions : Marché > 50 000 € HT  ET  durée > 2 mois
  Si durée ≤ 12 mois : Avance = 10% × Montant initial TTC
  Si durée > 12 mois : Avance = 10% × (12 × Montant TTC / N)

REMBOURSEMENT AVANCE
  Début remboursement : avancement cumulé ≥ 65%
  Fin remboursement   : avancement cumulé = 100%
  R_cumulé = Avance × (% cum. - 65%) / 35%
  R_mois   = R_cumulé - Remboursements déjà effectués

RETENUE DE GARANTIE
  RG mensuelle = 5% × Montant TTC du décompte du mois
  Plafond = 5% × Montant initial TTC du marché

STRUCTURE DU DÉCOMPTE MENSUEL
  Travaux HT   = Montant HT × (% cum. M - % cum. M-1)
  Révision HT  = Travaux HT × (Cn - 1)
  Montant HT net = Travaux HT + Révision HT - Remboursement avance
  TVA          = Montant HT net × 20%
  Montant TTC  = Montant HT net + TVA
  RG           = Montant TTC × 5%
  Net à payer  = Montant TTC - RG - Acomptes antérieurs
```

---

## EXERCICE 1 — Calcul de l'avance forfaitaire (type sujet 2023)

### Énoncé
Marché d'une entreprise de maçonnerie :
- Montant HT du marché : **342 500 € HT**
- TVA : 20%
- Durée du marché : **9 mois**

**Questions :**
1. L'entreprise a-t-elle droit à une avance forfaitaire ? Justifier.
2. Calculer le montant TTC du marché.
3. Calculer le montant de l'avance forfaitaire.

---

### CORRIGÉ

**1. Conditions d'éligibilité :**

```
Condition 1 : Montant HT > 50 000 € → 342 500 > 50 000 ✓
Condition 2 : Durée > 2 mois → 9 mois > 2 mois ✓

→ L'entreprise a DROIT à l'avance forfaitaire.
```

**2. Montant TTC :**

```
TVA = 342 500 × 20% = 68 500 €
Montant TTC = 342 500 + 68 500 = 411 000 € TTC
```

**3. Avance forfaitaire (durée ≤ 12 mois) :**

```
Avance = 10% × Montant TTC
Avance = 10% × 411 000
Avance = 41 100 € TTC
```

---

## EXERCICE 2 — Remboursement de l'avance (type sujet 2023-2024)

### Énoncé
Reprise de l'exercice 1 : Avance = **41 100 € TTC**

Tableau d'avancement :

| Mois | % cumulé |
|---|---|
| M1 | 8% |
| M2 | 18% |
| M3 | 32% |
| M4 | 50% |
| M5 | 68% |
| M6 | 82% |
| M7 | 92% |
| M8 | 100% |

**Questions :**
1. À partir de quel mois commence le remboursement ?
2. Calculer R_cumulé et R_mois pour les mois M5 à M8.

---

### CORRIGÉ

**1. Début du remboursement :**

```
Le remboursement commence quand le % cumulé ≥ 65%.
M4 = 50% < 65%  →  pas de remboursement
M5 = 68% ≥ 65%  →  DÉBUT du remboursement en M5
```

**2. Tableau de remboursement :**

```
Formule : R_cumulé = Avance × (% cum. - 65%) / 35%
          R_mois   = R_cumulé - Remb. déjà effectués

MOIS M5 (68%) :
  R_cumulé = 41 100 × (68% - 65%) / 35%
  R_cumulé = 41 100 × 3% / 35%
  R_cumulé = 41 100 × 0,08571
  R_cumulé = 3 522,78 €
  R_mois M5 = 3 522,78 - 0 = 3 522,78 €

MOIS M6 (82%) :
  R_cumulé = 41 100 × (82% - 65%) / 35%
  R_cumulé = 41 100 × 17% / 35%
  R_cumulé = 41 100 × 0,48571
  R_cumulé = 19 962,86 €
  R_mois M6 = 19 962,86 - 3 522,78 = 16 440,08 €

MOIS M7 (92%) :
  R_cumulé = 41 100 × (92% - 65%) / 35%
  R_cumulé = 41 100 × 27% / 35%
  R_cumulé = 41 100 × 0,77143
  R_cumulé = 31 705,71 €
  R_mois M7 = 31 705,71 - 19 962,86 = 11 742,85 €

MOIS M8 (100%) :
  R_cumulé = 41 100 × (100% - 65%) / 35%
  R_cumulé = 41 100 × 35% / 35%
  R_cumulé = 41 100 × 1,00
  R_cumulé = 41 100,00 €  (= montant total de l'avance)
  R_mois M8 = 41 100,00 - 31 705,71 = 9 394,29 €
```

**Vérification :**
```
Total remboursé = 3 522,78 + 16 440,08 + 11 742,85 + 9 394,29
               = 41 100,00 €  ✓  (= avance totale remboursée)
```

---

## EXERCICE 3 — Décompte mensuel complet (type sujet 2025 — Halle de Maubeuge)

### Énoncé
Marché de charpente métallique :
- Montant initial HT : **280 000 € HT** — TVA 20%
- Durée : 8 mois — Avance : **33 600 € TTC** (calculée précédemment)
- Coefficient de révision Cn mois M4 : **1,043**

Avancement :

| Mois | % cumulé |
|---|---|
| M1 | 5% |
| M2 | 15% |
| M3 | 35% |
| M4 | 60% |
| M5 | 78% |
| M6 | 90% |
| M7 | 97% |
| M8 | 100% |

Établir le **décompte complet du mois M4** (aucun remboursement d'avance car 60% < 65%).

---

### CORRIGÉ

**Montant TTC du marché :**
```
280 000 × 1,20 = 336 000 € TTC
```

**Décompte mois M4 :**

```
1. TRAVAUX HT DU MOIS
   Travaux HT = Montant HT × (% cum. M4 - % cum. M3)
   Travaux HT = 280 000 × (60% - 35%)
   Travaux HT = 280 000 × 25%
   Travaux HT = 70 000 € HT

2. RÉVISION HT
   Révision HT = Travaux HT × (Cn - 1)
   Révision HT = 70 000 × (1,043 - 1)
   Révision HT = 70 000 × 0,043
   Révision HT = 3 010 € HT

3. REMBOURSEMENT AVANCE
   % cumulé M4 = 60% < 65%
   → Pas de remboursement d'avance ce mois.
   Remb. avance = 0 €

4. MONTANT HT NET
   Montant HT net = Travaux HT + Révision HT - Remb. avance
   Montant HT net = 70 000 + 3 010 - 0
   Montant HT net = 73 010 € HT

5. TVA
   TVA = 73 010 × 20% = 14 602 €

6. MONTANT TTC
   Montant TTC = 73 010 + 14 602 = 87 612 € TTC

7. RETENUE DE GARANTIE
   RG = 87 612 × 5% = 4 380,60 €
   Plafond RG = 336 000 × 5% = 16 800 €
   RG cumulée vérification : à calculer sur les mois précédents
   → RG mois M4 = 4 380,60 € (plafond non atteint)

8. NET À PAYER (hors acomptes antérieurs)
   Net à payer = Montant TTC - RG
   Net à payer = 87 612 - 4 380,60
   Net à payer = 83 231,40 € TTC
```

**Tableau synthèse décompte M4 :**

| Poste | Montant |
|---|---|
| Travaux HT | 70 000,00 € |
| Révision HT | 3 010,00 € |
| Remboursement avance | 0,00 € |
| **Montant HT net** | **73 010,00 €** |
| TVA 20% | 14 602,00 € |
| **Montant TTC** | **87 612,00 €** |
| Retenue de garantie (5%) | -4 380,60 € |
| **Net à payer TTC** | **83 231,40 €** |

---

## EXERCICE 4 — Décompte mois M5 avec remboursement avance

### Énoncé
Suite de l'exercice 3. Au mois M5 : % cumulé = 78%, Cn = 1,051.

Établir le décompte complet du mois M5.

---

### CORRIGÉ

```
1. TRAVAUX HT
   Travaux HT = 280 000 × (78% - 60%) = 280 000 × 18% = 50 400 € HT

2. RÉVISION HT
   Révision HT = 50 400 × (1,051 - 1) = 50 400 × 0,051 = 2 570,40 € HT

3. REMBOURSEMENT AVANCE
   R_cumulé M5 = 33 600 × (78% - 65%) / 35%
   R_cumulé M5 = 33 600 × 13% / 35%
   R_cumulé M5 = 33 600 × 0,37143
   R_cumulé M5 = 12 480,00 €
   Remb. M4    = 0 € (pas de remboursement en M4 car 60% < 65%)
   R_mois M5   = 12 480,00 - 0 = 12 480,00 €

4. MONTANT HT NET
   Montant HT net = 50 400 + 2 570,40 - 12 480
   Montant HT net = 40 490,40 € HT

5. TVA = 40 490,40 × 20% = 8 098,08 €

6. MONTANT TTC = 40 490,40 + 8 098,08 = 48 588,48 € TTC

7. RETENUE DE GARANTIE = 48 588,48 × 5% = 2 429,42 €

8. NET À PAYER = 48 588,48 - 2 429,42 = 46 159,06 € TTC
```

---

## POINTS DE VIGILANCE

- La **RG s'applique sur le TTC** (pas sur le HT)
- **Plafond RG** = 5% du montant initial TTC — vérifier à chaque mois
- Le remboursement d'avance se déduit **avant** le calcul de TVA (il réduit le HT net)
- **Pas de révision sur l'avance** : la révision ne s'applique qu'aux travaux
- Le seuil de remboursement est **65%** (et non 60% ou 70%)
- Le dénominateur est toujours **35%** (= 100% - 65%)
